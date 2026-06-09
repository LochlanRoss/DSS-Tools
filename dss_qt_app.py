from __future__ import annotations

import csv
import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
from fnmatch import fnmatch
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable, Iterable

from openpyxl import Workbook

import dss_hours_tracker as core

from PySide6.QtCore import QAbstractTableModel, QModelIndex, QObject, QPoint, Qt, QTimer, QUrl, Signal
from PySide6.QtGui import QAction, QColor, QDesktopServices, QIcon, QKeySequence
from PySide6.QtWidgets import (
    QApplication,
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
    QSpinBox,
    QSplitter,
    QTabWidget,
    QTableView,
    QTextEdit,
    QToolButton,
    QVBoxLayout,
    QWidget,
)


TABLE_DATE_COLUMNS = {"date", "work_date", "week_start", "week_end"}
TABLE_NUMERIC_COLUMNS = {"st", "ot", "dt", "total", "expanded", "days", "similarity", "limit", "actual_total", "delta"}


@dataclass(frozen=True)
class TableSpec:
    table_id: str
    title: str
    columns: tuple[tuple[str, str], ...]


class TableModel(QAbstractTableModel):
    def __init__(self, columns: list[tuple[str, str]], rows: list[dict[str, Any]] | None = None, theme: core.UiThemeColors | None = None) -> None:
        super().__init__()
        self.columns = columns
        self.rows: list[dict[str, Any]] = rows or []
        self.theme = theme or core.DEFAULT_UI_THEME

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
            if "missing_email" in tags:
                return QColor(core.MISSING_EMAIL_ROW_FOREGROUND)
            if "alert" in tags:
                return QColor(self.theme.alert_row_foreground)
            if "crew_total" in tags:
                return QColor(self.theme.crew_total_foreground)
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
            return str(value).casefold()

        self.layoutAboutToBeChanged.emit()
        self.rows.sort(key=sort_key, reverse=reverse)
        self.layoutChanged.emit()


class DataTablePage(QWidget):
    layoutChanged = Signal(str, list, dict, str, bool)

    def __init__(self, spec: TableSpec, theme: core.UiThemeColors, config_path: Path) -> None:
        super().__init__()
        self.spec = spec
        self.config_path = config_path
        self.columns = list(spec.columns)
        self.model = TableModel(self.columns, theme=theme)
        self.view = QTableView(self)
        self.view.setModel(self.model)
        self.view.setSortingEnabled(True)
        self.view.setAlternatingRowColors(False)
        self.view.setSelectionBehavior(QTableView.SelectRows)
        self.view.setSelectionMode(QTableView.ExtendedSelection)
        self.view.verticalHeader().setVisible(False)
        self.view.horizontalHeader().setSectionsMovable(True)
        self.view.horizontalHeader().setStretchLastSection(False)
        self.view.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)

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
        self.view.horizontalHeader().sortIndicatorChanged.connect(self._emit_layout_change)
        self.view.horizontalHeader().sectionMoved.connect(lambda *_: self._emit_layout_change())
        self.view.horizontalHeader().sectionResized.connect(lambda *_: self._emit_layout_change())

    def add_toolbar_button(self, text: str, callback: Callable[[], None]) -> QPushButton:
        button = QPushButton(text, self)
        button.clicked.connect(callback)
        self.top_bar.insertWidget(0, button)
        return button

    def set_theme(self, theme: core.UiThemeColors) -> None:
        self.model.theme = theme
        self.model.layoutChanged.emit()

    def set_rows(self, rows: list[dict[str, Any]]) -> None:
        self._current_rows = rows
        self.model.set_rows(rows)

    def selected_rows(self) -> list[dict[str, Any]]:
        indexes = self.view.selectionModel().selectedRows()
        results: list[dict[str, Any]] = []
        for index in indexes:
            row = self.model.rows[index.row()]
            results.append(row)
        return results

    def export_rows(self) -> tuple[list[str], list[dict[str, Any]]]:
        visible_columns = self.visible_columns()
        return visible_columns, list(self.model.rows)

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


class CheckListPopup(QFrame):
    selectionChanged = Signal(list)

    def __init__(self, parent: QWidget, all_label: str, clear_label: str) -> None:
        super().__init__(parent, Qt.Popup | Qt.FramelessWindowHint)
        self.all_label = all_label
        self.clear_label = clear_label
        self.values: list[str] = []
        self.selected: set[str] = set()
        self.list_widget = QListWidget(self)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.addWidget(self.list_widget)
        self.list_widget.itemChanged.connect(self._on_item_changed)

    def set_values(self, values: list[str], selected: Iterable[str]) -> None:
        self.values = list(values)
        self.selected = set(selected)
        self.list_widget.blockSignals(True)
        self.list_widget.clear()
        for label, data, checked in [
            (self.all_label, "__all__", len(self.selected) == len(self.values) and len(self.values) > 0),
            (self.clear_label, "__clear__", False),
        ]:
            item = QListWidgetItem(label)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setData(Qt.UserRole, data)
            item.setCheckState(Qt.Checked if checked else Qt.Unchecked)
            self.list_widget.addItem(item)
        for value in self.values:
            item = QListWidgetItem(value)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setData(Qt.UserRole, value)
            item.setCheckState(Qt.Checked if value in self.selected else Qt.Unchecked)
            self.list_widget.addItem(item)
        self.list_widget.blockSignals(False)

    def _on_item_changed(self, item: QListWidgetItem) -> None:
        data = item.data(Qt.UserRole)
        if data == "__all__":
            selected = item.checkState() == Qt.Checked
            self.list_widget.blockSignals(True)
            for row in range(2, self.list_widget.count()):
                self.list_widget.item(row).setCheckState(Qt.Checked if selected else Qt.Unchecked)
            self.list_widget.blockSignals(False)
        elif data == "__clear__":
            self.list_widget.blockSignals(True)
            item.setCheckState(Qt.Unchecked)
            for row in range(2, self.list_widget.count()):
                self.list_widget.item(row).setCheckState(Qt.Unchecked)
            self.list_widget.blockSignals(False)
        selected_values: list[str] = []
        for row in range(2, self.list_widget.count()):
            entry = self.list_widget.item(row)
            if entry.checkState() == Qt.Checked:
                selected_values.append(str(entry.data(Qt.UserRole)))
        self.selected = set(selected_values)
        self.selectionChanged.emit(selected_values)


class CheckListButton(QPushButton):
    selectionChanged = Signal(list)

    def __init__(self, all_text: str, noun: str, parent: QWidget | None = None) -> None:
        super().__init__(all_text, parent)
        self.all_text = all_text
        self.noun = noun
        self.values: list[str] = []
        self.selected: list[str] = []
        self.popup = CheckListPopup(self, all_text, "Uncheck All")
        self.popup.selectionChanged.connect(self._set_selection)
        self.clicked.connect(self._show_popup)

    def set_choices(self, values: Iterable[str], selected: Iterable[str] | None = None, force_single: bool = False) -> None:
        self.values = list(values)
        self.setEnabled(bool(self.values))
        if force_single and len(self.values) == 1:
            self.selected = [self.values[0]]
            self.setEnabled(False)
        else:
            allowed = set(self.values)
            desired = [value for value in (selected or self.selected or self.values) if value in allowed]
            self.selected = desired if desired else list(self.values)
        self._update_text()

    def selected_values(self) -> list[str]:
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
    def __init__(self, parent: QWidget, root_folder: Path, selected_paths: Iterable[Path] | None = None, title: str = "Quick DSS Picker") -> None:
        super().__init__(parent)
        self.root_folder = root_folder
        self.selected_paths = {str(Path(path).resolve()) for path in (selected_paths or [])}
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

        self._populate()

    def _iter_candidate_paths(self) -> list[Path]:
        candidates: list[Path] = []
        if not self.root_folder.exists() or not self.root_folder.is_dir():
            return candidates
        for job_dir in sorted((path for path in self.root_folder.iterdir() if path.is_dir()), key=lambda p: p.name.casefold()):
            try:
                subdirs = [child for child in job_dir.iterdir() if child.is_dir() and "dss" in child.name.casefold()]
            except OSError:
                continue
            for dss_dir in sorted(subdirs, key=lambda p: p.name.casefold()):
                try:
                    files = [child for child in dss_dir.iterdir() if child.is_file() and child.suffix.lower() == ".xlsx"]
                except OSError:
                    continue
                for file_path in files:
                    if file_path.name.startswith("~$"):
                        continue
                    candidates.append(file_path.resolve())
        try:
            candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        except OSError:
            candidates.sort(key=lambda p: p.name.casefold())
        return candidates

    def _format_item_label(self, path: Path) -> str:
        pf = core.extract_pf_identifier(path.name)
        job_folder = path.parent.parent.name if path.parent.parent else ""
        try:
            modified = datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
        except OSError:
            modified = "Unknown"
        return f"{pf} | {job_folder} | {path.name} | {modified}\n{path}"

    def _populate(self) -> None:
        previous = set(self.selected_file_paths()) | set(self.selected_paths)
        self.list_widget.clear()
        candidates = self._iter_candidate_paths()
        for path in candidates:
            item = QListWidgetItem(self._format_item_label(path))
            item.setData(Qt.UserRole, str(path))
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked if str(path) in previous else Qt.Unchecked)
            self.list_widget.addItem(item)
        self.summary_label.setText(f"Found {len(candidates)} DSS workbook(s) under {self.root_folder}")
        self._apply_filter()

    def _apply_filter(self) -> None:
        needle = self.search_edit.text().strip().casefold()
        visible = 0
        for row in range(self.list_widget.count()):
            item = self.list_widget.item(row)
            haystack = item.text().casefold()
            hidden = bool(needle) and needle not in haystack
            item.setHidden(hidden)
            if not hidden:
                visible += 1
        self.summary_label.setText(f"Showing {visible} DSS workbook(s) from {self.root_folder}")

    def _check_all_visible(self) -> None:
        for row in range(self.list_widget.count()):
            item = self.list_widget.item(row)
            if not item.isHidden():
                item.setCheckState(Qt.Checked)

    def _uncheck_all_visible(self) -> None:
        for row in range(self.list_widget.count()):
            item = self.list_widget.item(row)
            if not item.isHidden():
                item.setCheckState(Qt.Unchecked)

    def selected_file_paths(self) -> list[str]:
        results: list[str] = []
        for row in range(self.list_widget.count()):
            item = self.list_widget.item(row)
            if item.checkState() == Qt.Checked:
                results.append(str(item.data(Qt.UserRole)))
        return results


class LoadWorker(QObject):
    progressChanged = Signal(float, str)
    partialReady = Signal(object, str)
    finished = Signal(object)
    failed = Signal(str)
    cancelled = Signal()

    def __init__(self, source_paths: list[Path], previous_data: core.TrackerData | None, cache_dir: Path) -> None:
        super().__init__()
        self.source_paths = source_paths
        self.previous_data = previous_data
        self.cache_dir = cache_dir
        self.cancel_event = threading.Event()

    def cancel(self) -> None:
        self.cancel_event.set()

    def run(self) -> None:
        try:
            data = core.load_tracker_data(
                self.source_paths,
                previous_data=self.previous_data,
                progress_callback=lambda pct, msg: self.progressChanged.emit(float(pct), str(msg)),
                partial_callback=lambda snapshot, msg: self.partialReady.emit(snapshot, msg),
                cache_dir=self.cache_dir,
                should_cancel=self.cancel_event.is_set,
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

    def __init__(self, employee_names: list[str]) -> None:
        super().__init__()
        self.employee_names = employee_names
        self.cancel_event = threading.Event()

    def cancel(self) -> None:
        self.cancel_event.set()

    def run(self) -> None:
        try:
            results, address_book_names = core.query_outlook_emails(self.employee_names, should_cancel=self.cancel_event.is_set)
        except core.OperationCancelled:
            self.cancelled.emit()
            return
        except Exception as exc:  # pragma: no cover - UI surface
            self.failed.emit(str(exc))
            return
        self.finished.emit(results, address_book_names)


class EmployeesPage(QWidget):
    changed = Signal()
    syncRequested = Signal()

    def __init__(self) -> None:
        super().__init__()
        self.employee_names: list[str] = []
        self.employee_emails: dict[str, str] = {}
        self.employee_notes: dict[str, str] = {}
        self.employee_groups: dict[str, list[str]] = {}
        self.missing_email_suppressions: set[str] = set()

        splitter = QSplitter(self)
        left = QWidget()
        left_layout = QVBoxLayout(left)
        self.employee_list = QListWidget()
        left_buttons = QHBoxLayout()
        self.add_employee_button = QPushButton("Add Employee")
        self.remove_employee_button = QPushButton("Remove Employee")
        self.sync_button = QPushButton("Sync Outlook Emails")
        left_buttons.addWidget(self.add_employee_button)
        left_buttons.addWidget(self.remove_employee_button)
        left_buttons.addWidget(self.sync_button)
        left_layout.addWidget(QLabel("Employees"))
        left_layout.addWidget(self.employee_list, 1)
        left_layout.addLayout(left_buttons)

        right = QWidget()
        right_layout = QVBoxLayout(right)
        detail_box = QGroupBox("Employee Details")
        form = QFormLayout(detail_box)
        self.employee_name_label = QLabel("")
        self.email_edit = QLineEdit()
        self.suppression_box = QCheckBox("Suppress missing email warnings for this employee")
        self.notes_edit = QPlainTextEdit()
        self.group_list = QListWidget()
        self.group_list.setSelectionMode(QListWidget.NoSelection)
        form.addRow("Employee", self.employee_name_label)
        form.addRow("Email", self.email_edit)
        form.addRow("", self.suppression_box)
        form.addRow("Notes", self.notes_edit)
        form.addRow("Groups", self.group_list)

        group_box = QGroupBox("Groups")
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
        self.groups_list.currentItemChanged.connect(self._group_selected)
        self.add_employee_button.clicked.connect(self._add_employee)
        self.remove_employee_button.clicked.connect(self._remove_employee)
        self.add_group_button.clicked.connect(self._add_group)
        self.remove_group_button.clicked.connect(self._remove_group)
        self.sync_button.clicked.connect(self.syncRequested.emit)
        self.email_edit.editingFinished.connect(self._save_current)
        self.suppression_box.toggled.connect(lambda _checked: self._save_current())
        self.notes_edit.textChanged.connect(self._save_current)
        self.group_list.itemChanged.connect(lambda _item: self._save_current())

    def set_data(
        self,
        employee_names: list[str],
        employee_emails: dict[str, str],
        employee_notes: dict[str, str],
        employee_groups: dict[str, list[str]],
        missing_email_suppressions: set[str],
    ) -> None:
        current = self.current_employee()
        self.employee_names = list(employee_names)
        self.employee_emails = dict(employee_emails)
        self.employee_notes = dict(employee_notes)
        self.employee_groups = {name: list(values) for name, values in employee_groups.items()}
        self.missing_email_suppressions = set(missing_email_suppressions)

        self.employee_list.blockSignals(True)
        self.employee_list.clear()
        for employee in self.employee_names:
            suppressed = employee in self.missing_email_suppressions
            label, missing = core.build_employee_email_list_label(employee, self.employee_emails.get(employee, ""), suppressed=suppressed)
            item = QListWidgetItem(label)
            item.setData(Qt.UserRole, employee)
            if missing:
                item.setBackground(QColor(core.MISSING_EMAIL_ROW_BACKGROUND))
                item.setForeground(QColor(core.MISSING_EMAIL_ROW_FOREGROUND))
            self.employee_list.addItem(item)
        self.employee_list.blockSignals(False)

        self.groups_list.blockSignals(True)
        self.groups_list.clear()
        for group_name in sorted(self.employee_groups, key=str.casefold):
            item = QListWidgetItem(group_name)
            item.setData(Qt.UserRole, group_name)
            self.groups_list.addItem(item)
        self.groups_list.blockSignals(False)

        if current:
            matches = self.employee_list.findItems("", Qt.MatchContains)
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

    def snapshot(self) -> tuple[list[str], dict[str, str], dict[str, str], dict[str, list[str]], set[str]]:
        return (
            list(self.employee_names),
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
        self.changed.emit()

    def _remove_employee(self) -> None:
        employee = self.current_employee()
        if not employee:
            return
        if QMessageBox.question(self, "Remove Employee", f"Hide '{employee}' from the managed roster?") != QMessageBox.Yes:
            return
        if employee in self.employee_names:
            self.employee_names.remove(employee)
        self.employee_emails.pop(employee, None)
        self.employee_notes.pop(employee, None)
        self.missing_email_suppressions.discard(employee)
        for group_name, members in list(self.employee_groups.items()):
            self.employee_groups[group_name] = [member for member in members if member != employee]
        self.changed.emit()

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


class FormattingRulesPage(QWidget):
    changed = Signal()

    def __init__(self) -> None:
        super().__init__()
        self.profiles: dict[str, core.FormattingProfile] = {}
        self.current_profile_name = core.DEFAULT_PROFILE_NAME
        self.job_presets: dict[str, str] = {}

        layout = QVBoxLayout(self)
        form_box = QGroupBox("Formatting Rules")
        form = QFormLayout(form_box)
        self.profile_combo = QComboBox()
        self.daily_st_edit = QLineEdit()
        self.weekly_st_edit = QLineEdit()
        self.weekly_ot_edit = QLineEdit()
        self.max_hours_edit = QLineEdit()
        self.job_preset_combo = QComboBox()
        self.job_preset_combo.setEditable(True)
        form.addRow("Profile", self.profile_combo)
        form.addRow("Job Preset", self.job_preset_combo)
        form.addRow("Daily ST Alert", self.daily_st_edit)
        form.addRow("Weekly ST Alert", self.weekly_st_edit)
        form.addRow("Weekly OT Alert", self.weekly_ot_edit)
        form.addRow("Max Hours Per Day", self.max_hours_edit)
        buttons = QHBoxLayout()
        self.new_button = QPushButton("New Profile")
        self.delete_button = QPushButton("Delete Profile")
        self.save_button = QPushButton("Save Profile")
        buttons.addWidget(self.new_button)
        buttons.addWidget(self.delete_button)
        buttons.addWidget(self.save_button)

        presets_box = QGroupBox("Rule Presets by Job")
        presets_layout = QVBoxLayout(presets_box)
        self.presets_list = QListWidget()
        preset_buttons = QHBoxLayout()
        self.apply_preset_button = QPushButton("Apply Job Preset")
        self.add_preset_button = QPushButton("Save Job Preset")
        self.remove_preset_button = QPushButton("Delete Job Preset")
        preset_buttons.addWidget(self.apply_preset_button)
        preset_buttons.addWidget(self.add_preset_button)
        preset_buttons.addWidget(self.remove_preset_button)
        presets_layout.addWidget(self.presets_list, 1)
        presets_layout.addLayout(preset_buttons)

        layout.addWidget(form_box)
        layout.addLayout(buttons)
        layout.addWidget(presets_box, 1)

        self.profile_combo.currentTextChanged.connect(self._profile_changed)
        self.presets_list.currentItemChanged.connect(self._preset_selected)
        self.new_button.clicked.connect(self._new_profile)
        self.delete_button.clicked.connect(self._delete_profile)
        self.save_button.clicked.connect(self._save_profile)
        self.apply_preset_button.clicked.connect(self._apply_preset)
        self.add_preset_button.clicked.connect(self._add_preset)
        self.remove_preset_button.clicked.connect(self._remove_preset)

    def set_data(self, profiles: dict[str, core.FormattingProfile], current_profile_name: str, job_presets: dict[str, str]) -> None:
        self.profiles = dict(profiles)
        self.current_profile_name = current_profile_name
        self.job_presets = dict(job_presets)
        self.profile_combo.blockSignals(True)
        self.profile_combo.clear()
        for name in sorted(self.profiles, key=str.casefold):
            self.profile_combo.addItem(name)
        self.profile_combo.setCurrentText(current_profile_name)
        self.profile_combo.blockSignals(False)
        self._load_profile(self.profiles[self.current_profile_name])
        self._refresh_presets()

    def snapshot(self) -> tuple[dict[str, core.FormattingProfile], str, dict[str, str]]:
        return dict(self.profiles), self.current_profile_name, dict(self.job_presets)

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

    def _profile_from_form(self, name: str) -> core.FormattingProfile:
        return core.FormattingProfile(
            name=name,
            st_threshold=core.parse_threshold_value(self.weekly_st_edit.text()) if self.weekly_st_edit.text().strip() else None,
            ot_threshold=core.parse_threshold_value(self.weekly_ot_edit.text()) if self.weekly_ot_edit.text().strip() else None,
            daily_st_threshold=core.parse_threshold_value(self.daily_st_edit.text()) if self.daily_st_edit.text().strip() else None,
            max_hours_per_day=core.parse_threshold_value(self.max_hours_edit.text()) if self.max_hours_edit.text().strip() else None,
        )

    def _new_profile(self) -> None:
        name, ok = QInputDialog.getText(self, "New Profile", "Profile name:")
        profile_name = name.strip()
        if not ok or not profile_name:
            return
        self.profiles[profile_name] = core.FormattingProfile(profile_name, None, None, None, None)
        self.current_profile_name = profile_name
        self.set_data(self.profiles, self.current_profile_name, self.job_presets)
        self.changed.emit()

    def _delete_profile(self) -> None:
        if len(self.profiles) <= 1:
            return
        name = self.profile_combo.currentText()
        if QMessageBox.question(self, "Delete Profile", f"Delete profile '{name}'?") != QMessageBox.Yes:
            return
        self.profiles.pop(name, None)
        self.current_profile_name = sorted(self.profiles, key=str.casefold)[0]
        self.set_data(self.profiles, self.current_profile_name, self.job_presets)
        self.changed.emit()

    def _save_profile(self) -> None:
        name = self.profile_combo.currentText().strip()
        if not name:
            return
        self.profiles[name] = self._profile_from_form(name)
        self.current_profile_name = name
        self.changed.emit()

    def _refresh_presets(self) -> None:
        current_job = self.job_preset_combo.currentText().strip()
        self.job_preset_combo.blockSignals(True)
        self.job_preset_combo.clear()
        self.job_preset_combo.addItems(sorted(self.job_presets, key=str.casefold))
        if current_job:
            self.job_preset_combo.setCurrentText(current_job)
        self.job_preset_combo.blockSignals(False)
        self.presets_list.clear()
        for job_name, profile_name in sorted(self.job_presets.items(), key=lambda item: item[0].casefold()):
            item = QListWidgetItem(f"{job_name} -> {profile_name}")
            item.setData(Qt.UserRole, job_name)
            self.presets_list.addItem(item)

    def _preset_selected(self, current: QListWidgetItem | None, _previous: QListWidgetItem | None) -> None:
        if current is None:
            return
        self.job_preset_combo.setCurrentText(str(current.data(Qt.UserRole)))

    def _apply_preset(self) -> None:
        job_name = self.job_preset_combo.currentText().strip()
        profile_name = self.job_presets.get(job_name, "")
        if not profile_name or profile_name not in self.profiles:
            QMessageBox.critical(self, "Formatting Rules", "That job preset is missing or points to a deleted profile.")
            return
        self.current_profile_name = profile_name
        self.profile_combo.setCurrentText(profile_name)
        self._load_profile(self.profiles[profile_name])
        self.changed.emit()

    def _add_preset(self) -> None:
        job_name = self.job_preset_combo.currentText().strip()
        if not job_name:
            QMessageBox.critical(self, "Formatting Rules", "Enter a job preset name first.")
            return
        self.job_presets[job_name] = self.profile_combo.currentText().strip() or self.current_profile_name
        self._refresh_presets()
        self.changed.emit()

    def _remove_preset(self) -> None:
        job_name = self.job_preset_combo.currentText().strip()
        if not job_name or job_name not in self.job_presets:
            return
        if QMessageBox.question(self, "Delete Job Preset", f"Delete job preset '{job_name}'?") != QMessageBox.Yes:
            return
        self.job_presets.pop(job_name, None)
        self._refresh_presets()
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

        settings_box = QGroupBox("Configuration")
        form = QFormLayout(settings_box)
        self.disable_typo_box = QCheckBox()
        self.show_daily_raw_box = QCheckBox()
        self.quickload_box = QCheckBox()
        self.hash_poll_spin = QSpinBox()
        self.hash_poll_spin.setRange(1, 1440)
        self.quickload_hotkey_combo = QComboBox()
        self.quickload_hotkey_combo.setEditable(True)
        self.quickload_hotkey_combo.addItems(list(core.QUICKLOAD_CANCEL_HOTKEY_PRESETS))
        self.auto_update_check_box = QCheckBox()
        self.auto_update_download_box = QCheckBox()
        self.library_root_edit = QLineEdit()
        self.library_root_browse_button = QPushButton("Browse...")
        self.version_label = QLabel(f"Application version: {core.APP_VERSION}")
        library_root_row = QWidget()
        library_root_layout = QHBoxLayout(library_root_row)
        library_root_layout.setContentsMargins(0, 0, 0, 0)
        library_root_layout.addWidget(self.library_root_edit, 1)
        library_root_layout.addWidget(self.library_root_browse_button)
        form.addRow("Disable name typo notifications", self.disable_typo_box)
        form.addRow("Show Daily Raw tab", self.show_daily_raw_box)
        form.addRow("Quick load last DSS set on startup", self.quickload_box)
        form.addRow("Check source DSS(s) frequency (minutes)", self.hash_poll_spin)
        form.addRow("Cancel quick-load hotkey", self.quickload_hotkey_combo)
        form.addRow("Automatically check GitHub for updates on startup", self.auto_update_check_box)
        form.addRow("Automatically download updates on unmetered Wi-Fi", self.auto_update_download_box)
        form.addRow("DSS library root folder", library_root_row)
        form.addRow("", self.version_label)

        appearance = QGroupBox("Appearance")
        appearance_layout = QGridLayout(appearance)
        self._theme_line_edits: dict[str, QLineEdit] = {}
        for row, (label, attr) in enumerate(core.UI_THEME_CONFIG_FIELDS):
            appearance_layout.addWidget(QLabel(label), row, 0)
            edit = QLineEdit()
            self._theme_line_edits[attr] = edit
            pick_button = QPushButton("Pick...")
            pick_button.clicked.connect(lambda _checked=False, key=attr: self._pick_theme_colour(key))
            appearance_layout.addWidget(edit, row, 1)
            appearance_layout.addWidget(pick_button, row, 2)
        self.reset_colours_button = QPushButton("Reset colours to sample defaults")
        self.reset_colours_button.clicked.connect(self._reset_theme_defaults)
        appearance_layout.addWidget(self.reset_colours_button, len(core.UI_THEME_CONFIG_FIELDS), 0, 1, 3)

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

        layout.addWidget(settings_box)
        layout.addWidget(appearance)
        layout.addWidget(self.apply_button, 0, Qt.AlignLeft)
        layout.addWidget(maintenance)
        layout.addWidget(diagnostics)
        layout.addStretch(1)

        self.disable_typo_box.toggled.connect(self.settingsChanged.emit)
        self.show_daily_raw_box.toggled.connect(self.settingsChanged.emit)
        self.quickload_box.toggled.connect(self.settingsChanged.emit)
        self.hash_poll_spin.valueChanged.connect(self.settingsChanged.emit)
        self.quickload_hotkey_combo.lineEdit().editingFinished.connect(self.settingsChanged.emit)
        self.auto_update_check_box.toggled.connect(self.settingsChanged.emit)
        self.auto_update_download_box.toggled.connect(self.settingsChanged.emit)
        self.library_root_edit.editingFinished.connect(self.settingsChanged.emit)
        for edit in self._theme_line_edits.values():
            edit.editingFinished.connect(self.settingsChanged.emit)
        self.library_root_browse_button.clicked.connect(self._browse_library_root)
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

    def _reset_theme_defaults(self) -> None:
        defaults = core.DEFAULT_UI_THEME
        for _label, attr in core.UI_THEME_CONFIG_FIELDS:
            edit = self._theme_line_edits.get(attr)
            if edit is not None:
                edit.setText(getattr(defaults, attr))
        self.settingsChanged.emit()

    def set_update_status(self, text: str) -> None:
        self.update_status_label.setText(text)

    def set_settings(self, settings: core.AppSettings) -> None:
        self.disable_typo_box.setChecked(settings.disable_name_typo_notifications)
        self.show_daily_raw_box.setChecked(settings.show_daily_raw_tab)
        self.quickload_box.setChecked(settings.quickload_last_sources_enabled)
        self.hash_poll_spin.setValue(settings.hash_poll_minutes)
        self.quickload_hotkey_combo.setCurrentText(settings.quickload_cancel_hotkey)
        self.auto_update_check_box.setChecked(settings.auto_update_check_enabled)
        self.auto_update_download_box.setChecked(settings.auto_download_updates_on_unmetered_wifi)
        self.library_root_edit.setText(settings.dss_library_root)
        for _label, attr in core.UI_THEME_CONFIG_FIELDS:
            edit = self._theme_line_edits.get(attr)
            if edit is not None:
                edit.setText(getattr(settings.ui_theme, attr))

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
            show_daily_raw_tab=self.show_daily_raw_box.isChecked(),
            quickload_last_sources_enabled=self.quickload_box.isChecked(),
            quickload_cancel_hotkey=hotkey,
            auto_update_check_enabled=self.auto_update_check_box.isChecked(),
            auto_download_updates_on_unmetered_wifi=self.auto_update_download_box.isChecked(),
            dss_library_root=self.library_root_edit.text().strip(),
            ui_theme=core.parse_ui_theme_payload(theme_payload, defaults=current.ui_theme),
        )


class EmailDraftsPage(QWidget):
    createDraftsRequested = Signal()
    syncRequested = Signal()
    templatesChanged = Signal()

    def __init__(self, theme: core.UiThemeColors, config_path: Path) -> None:
        super().__init__()
        self.preview_table = DataTablePage(
            TableSpec(
                "email_drafts",
                "Email Drafts",
                (
                    ("employee", "Employee"),
                    ("email", "Email"),
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

    def set_templates(self, subject_template: str, body_template: str) -> None:
        self.subject_edit.setPlainText(subject_template)
        self.body_edit.setPlainText(body_template)

    def set_weeks(self, week_starts: list[date]) -> None:
        current = self.selected_week_start()
        self.week_combo.blockSignals(True)
        self.week_combo.clear()
        for week_start in sorted(week_starts, reverse=True):
            self.week_combo.addItem(week_start.isoformat(), week_start)
        if current is not None:
            idx = self.week_combo.findText(current.isoformat())
            if idx >= 0:
                self.week_combo.setCurrentIndex(idx)
        self.week_combo.blockSignals(False)

    def selected_week_start(self) -> date | None:
        data = self.week_combo.currentData()
        return data if isinstance(data, date) else None


class DssQtMainWindow(QMainWindow):
    updateCheckResultReady = Signal(object, object, bool)
    updateCheckErrorRaised = Signal(str, bool)
    updateDownloadSuccessReady = Signal(object, object, bool, bool)
    updateDownloadErrorRaised = Signal(str, bool)

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
        self.employee_notes = core.load_employee_notes(self.config_path)
        self.employee_groups = core.load_employee_groups(self.config_path)
        self.missing_email_suppressions = core.load_missing_email_suppressions(self.config_path)
        self.employee_added_names, self.employee_hidden_names = core.load_employee_name_overrides(self.config_path)
        self.job_presets = core.load_job_presets(self.config_path)
        self.subject_template, self.body_template = core.load_email_templates(self.config_path)
        self.ignored_name_typos = core.load_ignored_name_typos(self.config_path)
        self.current_data: core.TrackerData | None = None
        self.source_paths: list[Path] = list(initial_source or [])
        self.load_worker: LoadWorker | None = None
        self.load_thread: threading.Thread | None = None
        self.outlook_worker: OutlookWorker | None = None
        self.outlook_thread: threading.Thread | None = None
        self._next_load_token = 0
        self._active_load_token = -1
        self._next_outlook_token = 0
        self._active_outlook_token = -1
        self._cancel_shortcut_action: QAction | None = None
        self._quickload_session = False
        self._update_check_in_progress = False
        self._update_download_in_progress = False
        self._auto_update_check_done = False
        self._downloaded_update_path: Path | None = None
        self._update_status_text = f"Installed version: {core.APP_VERSION}"
        self._hash_alerted_paths: set[Path] = set()
        self.hash_poll_timer = QTimer(self)
        self.updateCheckResultReady.connect(self._handle_update_check_result)
        self.updateCheckErrorRaised.connect(self._handle_update_check_error)
        self.updateDownloadSuccessReady.connect(self._handle_update_download_success)
        self.updateDownloadErrorRaised.connect(self._handle_update_download_error)
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
        QTimer.singleShot(core.AUTO_UPDATE_CHECK_DELAY_MS, self._auto_check_for_updates)

    def _build_ui(self) -> None:
        self.setWindowTitle(core.DISPLAY_APP_NAME)
        self.setMinimumSize(980, 640)
        root = QWidget()
        root_layout = QVBoxLayout(root)

        toolbar = QHBoxLayout()
        self.open_button = QPushButton("Open DSS Workbook(s)")
        self.quick_open_button = QPushButton("Quick Open")
        self.add_button = QPushButton("Add DSS")
        self.quick_add_button = QPushButton("Quick Add")
        self.remove_button = QPushButton("Remove DSS(s)")
        self.update_button = QPushButton("Update View")
        self.export_button = QPushButton("Export Current View")
        self.employee_filter = CheckListButton("All Employees", "Employees")
        self.pf_filter = CheckListButton("All PFs", "PFs")
        self.source_label = QLabel("No workbook loaded")
        self.loading_label = QLabel("")
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 1000)
        self.percent_label = QLabel("0.0%")
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.setEnabled(False)
        for widget in (
            self.open_button,
            self.quick_open_button,
            self.add_button,
            self.quick_add_button,
            self.remove_button,
            self.update_button,
            self.export_button,
            QLabel("Filter"),
            self.employee_filter,
            QLabel("PF"),
            self.pf_filter,
        ):
            if isinstance(widget, QWidget):
                toolbar.addWidget(widget)
        toolbar.addStretch(1)
        toolbar.addWidget(self.source_label, 1)
        toolbar.addWidget(self.loading_label)
        root_layout.addLayout(toolbar)

        status_row = QHBoxLayout()
        self.status_label = QLabel("Load a DSS workbook to view daily and weekly labour summaries.")
        self.status_label.setWordWrap(True)
        self.quickload_hint_label = QLabel("")
        self.quickload_hint_label.setWordWrap(True)
        status_row.addWidget(self.status_label, 1)
        status_row.addWidget(self.quickload_hint_label)
        status_row.addWidget(self.percent_label)
        status_row.addWidget(self.progress_bar)
        status_row.addWidget(self.cancel_button)
        root_layout.addLayout(status_row)

        self.group_tabs = QTabWidget()
        root_layout.addWidget(self.group_tabs, 1)
        self.setCentralWidget(root)

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
            (self.data_tabs, TableSpec("daily_raw", "Daily Raw", (("source_file", "Source File"), ("date", "Date"), ("sheet", "Source Sheet"), ("employee", "Employee"), ("st", "ST"), ("ot", "OT"), ("dt", "DT"), ("total", "Total"), ("expanded", "Expanded Hours"), ("ranges", "Source Ranges Used")))),
            (self.data_tabs, TableSpec("week_totals", "Week Totals", (("week_start", "Week Start"), ("week_end", "Week End"), ("st", "Whole Crew ST"), ("ot", "Whole Crew OT"), ("dt", "Whole Crew DT"), ("total", "Whole Crew Total"), ("expanded", "Expanded Hours")))),
            (self.summary_tabs, TableSpec("daily_by_pf", "Daily by PF", (("source_file", "Source File"), ("date", "Date"), ("employee", "Employee"), ("st", "ST"), ("ot", "OT"), ("dt", "DT"), ("total", "Total"), ("expanded", "Expanded Hours"), ("row_type", "Row Type")))),
            (self.summary_tabs, TableSpec("weekly_by_pf", "Weekly by PF", (("source_file", "Source File"), ("week_start", "Week Start"), ("week_end", "Week End"), ("employee", "Employee"), ("st", "ST"), ("ot", "OT"), ("dt", "DT"), ("total", "Total"), ("expanded", "Expanded Hours"), ("row_type", "Row Type")))),
            (self.summary_tabs, TableSpec("combined_daily", "Combined Summary Daily", (("date", "Date"), ("employee", "Employee"), ("st", "ST"), ("ot", "OT"), ("dt", "DT"), ("total", "Total"), ("expanded", "Expanded Hours")))),
            (self.summary_tabs, TableSpec("combined_weekly", "Combined Summary Weekly", (("week_start", "Week Start"), ("week_end", "Week End"), ("employee", "Employee"), ("st", "ST"), ("ot", "OT"), ("dt", "DT"), ("total", "Total"), ("expanded", "Expanded Hours")))),
            (self.report_tabs, TableSpec("error_report", "Error Report", (("employee", "Employee"), ("week_start", "Week Start"), ("week_end", "Week End"), ("hour_type", "Rule"), ("limit", "Limit"), ("actual_total", "Actual"), ("delta", "Delta"), ("trigger_date", "Trigger Date"), ("source_files", "Source Files"), ("reason", "Reason"), ("breakdown", "Breakdown")))),
            (self.report_tabs, TableSpec("parse_warnings", "Sheet Parse Warnings", (("source_file", "Source File"), ("sheet", "Sheet"), ("date", "Date"), ("issue", "Issue"), ("details", "Details")))),
            (self.report_tabs, TableSpec("workbook_health", "Workbook Health", (("source_file", "Source File"), ("status", "Status"), ("details", "Details")))),
            (self.report_tabs, TableSpec("audit_data_trail", "Audit Data Trail", (("source_file", "Source File"), ("date", "Date"), ("sheet", "Sheet"), ("employee", "Employee"), ("st", "ST"), ("ot", "OT"), ("dt", "DT"), ("total", "Total"), ("expanded", "Expanded Hours"), ("source_ranges", "Source Ranges"), ("audit", "Audit")))),
        ]:
            page = DataTablePage(spec, theme, self.config_path)
            page.layoutChanged.connect(self._save_table_layout)
            parent.addTab(page, spec.title)
            self.pages[spec.table_id] = page
        self.pages["error_report"].add_toolbar_button("Check Name Typos", self.check_name_typos_manually)

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

        self.open_button.clicked.connect(self.open_dss_files)
        self.quick_open_button.clicked.connect(self.quick_open_dss_files)
        self.add_button.clicked.connect(self.add_dss_files)
        self.quick_add_button.clicked.connect(self.quick_add_dss_files)
        self.remove_button.clicked.connect(self.remove_dss_files)
        self.update_button.clicked.connect(self.reload_data)
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
        self.formatting_page.set_data(self.profiles, self.current_profile_name, self.job_presets)
        self.email_drafts_page.set_templates(self.subject_template, self.body_template)
        self._sync_data_tabs_visibility()
        self.hash_poll_timer.setInterval(max(1, self.app_settings.hash_poll_minutes) * 60 * 1000)
        self._refresh_filters()
        self._refresh_employee_page()
        self._refresh_source_status_labels()
        self._refresh_quickload_hint_label()
        self._apply_ui_theme_chrome()
        self.refresh_views()
        self._set_loading_state(False)

    def _managed_employee_names(self) -> list[str]:
        discovered = set(self.current_data.employee_names if self.current_data else [])
        return sorted((discovered | self.employee_added_names) - self.employee_hidden_names, key=str.casefold)

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
        return sorted({core.extract_pf_identifier(path.name) for path in self.current_data.source_paths}, key=str.casefold)

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
        self.employees_page.set_data(
            self._managed_employee_names(),
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

    def _apply_ui_theme_chrome(self) -> None:
        theme = self.app_settings.ui_theme
        self.centralWidget().setStyleSheet(f"background-color: {theme.content_chrome_background};")

    def _sync_reports_alert_chrome(self, has_errors: bool, has_parse_warnings: bool) -> None:
        error_index = self.report_tabs.indexOf(self.pages["error_report"])
        parse_index = self.report_tabs.indexOf(self.pages["parse_warnings"])
        reports_index = self.group_tabs.indexOf(self.report_tabs)
        self.report_tabs.setTabText(error_index, "Error Report (!)" if has_errors else "Error Report")
        self.report_tabs.setTabText(parse_index, "Sheet Parse Warnings (!)" if has_parse_warnings else "Sheet Parse Warnings")
        self.group_tabs.setTabText(reports_index, "Reports (!)" if (has_errors or has_parse_warnings) else "Reports")

    def _active_profile(self) -> core.FormattingProfile:
        return self.profiles.get(self.current_profile_name, next(iter(self.profiles.values())))

    def _daily_records_filtered(self) -> list[core.DailyRecord]:
        if not self.current_data:
            return []
        allowed_employees = self._selected_employee_values()
        allowed_pfs = self._selected_pf_values()
        results: list[core.DailyRecord] = []
        for record in self.current_data.daily_records:
            if record.employee not in allowed_employees:
                continue
            if core.extract_pf_identifier(record.source_file) not in allowed_pfs:
                continue
            results.append(record)
        return results

    def refresh_views(self) -> None:
        for page in self.pages.values():
            page.set_theme(self.app_settings.ui_theme)
        self.email_drafts_page.preview_table.set_theme(self.app_settings.ui_theme)
        if not self.current_data:
            for page in self.pages.values():
                page.set_rows([])
            self.email_drafts_page.preview_table.set_rows([])
            self.email_drafts_page.set_weeks([])
            self._sync_reports_alert_chrome(False, False)
            return

        filtered_records = self._daily_records_filtered()
        profile = self._active_profile()
        daily_summary = core.aggregate_daily(filtered_records, combine_sources=False)
        weekly_summary = core.aggregate_weekly(filtered_records, combine_sources=False)
        daily_rollup = core.build_daily_rollup(daily_summary)
        weekly_rollup = core.build_weekly_rollup(weekly_summary)
        combined_daily = core.aggregate_daily(filtered_records, combine_sources=True)
        combined_weekly = core.aggregate_weekly(filtered_records, combine_sources=True)
        week_totals = core.build_week_totals(combined_weekly)
        findings = core.build_error_findings(filtered_records, profile)
        week_starts = sorted({record.week_start for record in combined_weekly}, reverse=True)

        self.email_drafts_page.set_weeks(week_starts)
        self.pages["daily_raw"].set_rows([
            {
                "source_file": record.source_file,
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
            for row in sorted(daily_rollup, key=lambda item: (item.work_date, item.source_file, item.row_type == "Crew Total", item.employee), reverse=True)
        ])
        self.pages["weekly_by_pf"].set_rows([
            {
                "source_file": row.source_file,
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
            for row in sorted(weekly_rollup, key=lambda item: (item.week_start, item.source_file, item.row_type == "Crew Total", item.employee), reverse=True)
        ])
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
        self.pages["error_report"].set_rows([
            {
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
                "__tags__": ("alert",),
            }
            for item in sorted(findings, key=lambda finding: (finding.week_start, finding.employee, finding.hour_type), reverse=True)
        ])
        filtered_source_files = {record.source_file for record in filtered_records}
        parse_warning_rows = [
            {
                "source_file": warning.source_file,
                "sheet": warning.source_sheet,
                "date": warning.work_date,
                "issue": warning.issue,
                "details": warning.details,
            }
            for warning in self.current_data.parse_warnings
            if warning.source_file in filtered_source_files
        ]
        self.pages["parse_warnings"].set_rows(parse_warning_rows)
        self.pages["workbook_health"].set_rows([
            {
                "source_file": item.source_file,
                "status": item.status,
                "details": item.details,
            }
            for item in self.current_data.workbook_health
            if item.source_file in {path.name for path in self.current_data.source_paths if core.extract_pf_identifier(path.name) in self._selected_pf_values()}
        ])
        self.pages["audit_data_trail"].set_rows([
            {
                "source_file": record.source_file,
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
            }
            for record in sorted(filtered_records, key=lambda item: (item.work_date, item.source_sheet, item.employee), reverse=True)
        ])
        self._refresh_email_preview(filtered_records)
        self._refresh_filters()
        self._sync_reports_alert_chrome(bool(findings), bool(parse_warning_rows))

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
            suppressed = request.employee in self.missing_email_suppressions
            display_email, missing = core.format_email_address_display(request.email, suppressed=suppressed)
            tags = ("missing_email",) if missing else ()
            rows.append(
                {
                    "employee": request.employee,
                    "email": display_email,
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
        employee_names, emails, notes, groups, suppressions = self.employees_page.snapshot()
        managed = set(employee_names)
        discovered = set(self.current_data.employee_names if self.current_data else [])
        self.employee_added_names = managed - discovered
        self.employee_hidden_names = discovered - managed
        self.employee_emails = emails
        self.employee_notes = notes
        self.employee_groups = groups
        self.missing_email_suppressions = suppressions
        core.save_employee_name_overrides(self.config_path, self.employee_added_names, self.employee_hidden_names)
        core.save_employee_emails(self.config_path, self.employee_emails, self.employee_outlook_display_names)
        core.save_employee_notes(self.config_path, self.employee_notes)
        core.save_employee_groups(self.config_path, self.employee_groups)
        core.save_missing_email_suppressions(self.config_path, self.missing_email_suppressions)
        self._refresh_filters()
        self.refresh_views()

    def _formatting_changed(self) -> None:
        self.profiles, self.current_profile_name, self.job_presets = self.formatting_page.snapshot()
        core.save_formatting_profiles(self.config_path, self.profiles, self.current_profile_name)
        core.save_job_presets(self.config_path, self.job_presets)
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
        self.refresh_views()

    def apply_settings(self) -> None:
        self._settings_changed()
        QMessageBox.information(self, "Configuration", "Settings applied.")

    def _set_loading_state(self, loading: bool, message: str = "") -> None:
        load_active = self.load_worker is not None
        outlook_active = self.outlook_worker is not None
        busy = load_active or outlook_active
        has_sources = bool(self.source_paths)
        has_data = self.current_data is not None
        self.open_button.setEnabled(not load_active)
        self.quick_open_button.setEnabled(not load_active)
        self.add_button.setEnabled(not load_active)
        self.quick_add_button.setEnabled(not load_active)
        self.remove_button.setEnabled(has_sources and not load_active)
        self.update_button.setEnabled(has_sources and not load_active)
        self.export_button.setEnabled(has_data and not busy)
        self.cancel_button.setEnabled(busy)
        self.loading_label.setText("Working..." if busy else "")
        self._refresh_quickload_hint_label()
        self._refresh_source_status_labels()
        if message:
            self.status_label.setText(message)

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

    def quick_open_dss_files(self) -> None:
        paths = self._show_quick_dss_picker("Quick Open DSS Workbooks")
        if not paths:
            return
        self.source_paths = paths
        core.save_last_open_dss_paths(self.config_path, self.source_paths)
        self.reload_data()

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

    def open_dss_files(self) -> None:
        paths, _filter = QFileDialog.getOpenFileNames(self, "Open DSS Workbook(s)", "", "Excel Workbook (*.xlsx)")
        if not paths:
            return
        self.source_paths = [Path(path).expanduser().resolve() for path in paths]
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
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(label)
        layout.addWidget(list_widget)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
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

    def reload_data(self) -> None:
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
        worker = LoadWorker(self.source_paths, self.current_data, self.cache_dir)
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
        self.loading_label.setText(f"{fraction * 100:.1f}%")
        self.status_label.setText(message)

    def _on_partial_ready(self, token: int, tracker_data: core.TrackerData, message: str) -> None:
        if token != self._active_load_token:
            return
        self.current_data = tracker_data
        self.status_label.setText(message)
        self._refresh_source_status_labels()
        self._refresh_filters()
        self._refresh_employee_page()
        if self.group_tabs.currentWidget() != self.settings_tabs:
            self.refresh_views()

    def _on_load_finished(self, token: int, tracker_data: core.TrackerData) -> None:
        if token != self._active_load_token:
            return
        self.current_data = tracker_data
        self.progress_bar.setValue(1000)
        self.percent_label.setText("100.0%")
        self.loading_label.setText("")
        self.load_worker = None
        self.load_thread = None
        self._quickload_session = False
        self._set_loading_state(False, f"Loaded {len(tracker_data.source_paths)} DSS workbook(s)")
        self._refresh_filters()
        self._refresh_employee_page()
        self.refresh_views()
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
        self.cancel_active_work(abandon_ui=True, reset_ui=False, message="")
        self._next_outlook_token += 1
        token = self._next_outlook_token
        self._active_outlook_token = token
        worker = OutlookWorker(employee_names)
        self.outlook_worker = worker
        worker.finished.connect(lambda results, address_book_names, current_token=token: self._on_outlook_sync_finished(current_token, results, address_book_names))
        worker.failed.connect(lambda message, current_token=token: self._on_outlook_sync_failed(current_token, message))
        worker.cancelled.connect(lambda current_token=token: self._on_outlook_sync_cancelled(current_token))
        self.outlook_thread = threading.Thread(target=worker.run, daemon=True)
        self._set_loading_state(True, "Syncing Outlook emails...")
        self.outlook_thread.start()

    def _on_outlook_sync_finished(self, token: int, results: dict[str, core.OutlookResolution], address_book_names: list[str]) -> None:
        if token != self._active_outlook_token:
            return
        self.outlook_worker = None
        self.outlook_thread = None
        updated = 0
        for employee, resolution in results.items():
            if resolution.email and not self.employee_emails.get(employee, "").strip():
                self.employee_emails[employee] = resolution.email
                if resolution.display_name.strip():
                    self.employee_outlook_display_names[employee] = resolution.display_name.strip()
                updated += 1
        core.save_employee_emails(self.config_path, self.employee_emails, self.employee_outlook_display_names)
        warnings = []
        if not self.app_settings.disable_name_typo_notifications:
            unresolved = [employee for employee in self._managed_employee_names() if not self.employee_emails.get(employee, "").strip() and employee not in self.missing_email_suppressions]
            warnings.extend(core.find_potential_name_typos(unresolved, self._managed_employee_names()))
            warnings.extend(core.find_address_book_name_typos(unresolved, address_book_names))
        if self.load_worker is None:
            self._set_loading_state(False, f"Matched emails: {updated}")
        self._refresh_employee_page()
        self.refresh_views()
        warnings = [
            warning
            for warning in warnings
            if core.typo_warning_key(warning.employee, warning.similar_employee) not in self.ignored_name_typos
        ]
        missing_after = sum(
            1
            for employee in self._managed_employee_names()
            if not self.employee_emails.get(employee, "").strip() and employee not in self.missing_email_suppressions
        )
        QMessageBox.information(self, "Outlook Email Sync", f"Matched emails: {updated}\nStill missing: {missing_after}")
        if warnings and not self.app_settings.disable_name_typo_notifications:
            self._show_typo_warning_dialog(warnings)

    def _on_outlook_sync_failed(self, token: int, message: str) -> None:
        if token != self._active_outlook_token:
            return
        self.outlook_worker = None
        self.outlook_thread = None
        if self.load_worker is None:
            self._set_loading_state(False, "Outlook sync failed")
        QMessageBox.critical(self, "Outlook Email Sync", message)

    def _on_outlook_sync_cancelled(self, token: int) -> None:
        if token != self._active_outlook_token:
            return
        self.outlook_worker = None
        self.outlook_thread = None
        if self.load_worker is None:
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
        core.remove_config_keys(self.config_path, ["employee_emails", "employee_outlook_display_names"])
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
        self.employee_notes = {}
        self.employee_groups = {}
        self.missing_email_suppressions = set()
        self.employee_added_names = set()
        self.employee_hidden_names = set()
        self.job_presets = {}
        self.subject_template, self.body_template = core.load_email_templates(self.config_path)
        self.ignored_name_typos = set()
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
        self.formatting_page.set_data(self.profiles, self.current_profile_name, self.job_presets)
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
                "show_daily_raw_tab": self.app_settings.show_daily_raw_tab,
                "quickload_last_sources_enabled": self.app_settings.quickload_last_sources_enabled,
                "quickload_cancel_hotkey": self.app_settings.quickload_cancel_hotkey,
                "auto_update_check_enabled": self.app_settings.auto_update_check_enabled,
                "auto_download_updates_on_unmetered_wifi": self.app_settings.auto_download_updates_on_unmetered_wifi,
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

    def check_name_typos_manually(self) -> None:
        if self.current_data is None or not self._managed_employee_names():
            QMessageBox.information(self, "Check Name Typos", "Load DSS data first.")
            return
        employee_names = self._managed_employee_names()
        daily_records = self.current_data.daily_records
        names_to_check = [
            employee
            for employee in employee_names
            if not self.employee_emails.get(employee, "").strip() and employee not in self.missing_email_suppressions
        ]
        raw_warnings: list[core.NameTypoWarning] = []
        address_book_names: list[str] = []
        if names_to_check:
            try:
                _results, address_book_names = core.query_outlook_emails(names_to_check)
            except Exception:
                address_book_names = []
        if names_to_check:
            raw_warnings.extend(core.find_potential_name_typos(names_to_check, employee_names, daily_records))
        else:
            raw_warnings.extend(core.find_similar_employee_name_pairs(employee_names, daily_records))
        if address_book_names:
            raw_warnings.extend(core.find_address_book_name_typos(names_to_check, address_book_names, daily_records))
        raw_warnings.extend(core.find_outlook_display_name_typos(employee_names, self.employee_outlook_display_names, daily_records))
        deduped: list[core.NameTypoWarning] = []
        seen_keys: set[str] = set()
        for warning in raw_warnings:
            key = core.typo_warning_key(warning.employee, warning.similar_employee)
            if key in seen_keys:
                continue
            seen_keys.add(key)
            deduped.append(warning)
        warnings = [
            warning
            for warning in deduped
            if core.typo_warning_key(warning.employee, warning.similar_employee) not in self.ignored_name_typos
        ]
        if not warnings:
            QMessageBox.information(self, "Check Name Typos", "No likely name typos were found.")
            return
        self._show_typo_warning_dialog(warnings)

    def _show_typo_warning_dialog(self, typo_warnings: list[core.NameTypoWarning]) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle("Potential Name Typos")
        dialog.resize(780, 420)
        layout = QVBoxLayout(dialog)
        intro = QLabel(
            "Possible name typo(s): similar names on the DSS roster, or a roster name that does not match the Outlook address book display name (after Sync Outlook Emails). Select an entry and choose Ignore Selected to stop warning on that pair. Notifications can also be disabled in Settings > Configuration."
        )
        intro.setWordWrap(True)
        layout.addWidget(intro)
        content = QHBoxLayout()
        list_widget = QListWidget()
        list_widget.setSelectionMode(QListWidget.ExtendedSelection)
        details = QTextEdit()
        details.setReadOnly(True)
        content.addWidget(list_widget, 2)
        content.addWidget(details, 3)
        layout.addLayout(content, 1)

        for warning in typo_warnings:
            item = QListWidgetItem(f"{warning.employee} -> {warning.similar_employee} ({warning.similarity * 100:.0f}% similar)")
            item.setData(Qt.UserRole, warning)
            list_widget.addItem(item)

        def refresh_details() -> None:
            item = list_widget.currentItem()
            if item is None:
                details.clear()
                return
            warning = item.data(Qt.UserRole)
            if not isinstance(warning, core.NameTypoWarning):
                details.clear()
                return
            lines = [
                f"{warning.employee} may match {warning.similar_employee}",
                f"Similarity: {warning.similarity * 100:.0f}%",
                "",
                "Locations:",
                *(warning.locations or ["(No locations found)"]),
            ]
            details.setPlainText("\n".join(lines))

        list_widget.currentItemChanged.connect(lambda _current, _previous: refresh_details())
        if list_widget.count():
            list_widget.setCurrentRow(0)

        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        ignore_button = QPushButton("Ignore Selected")
        buttons.addButton(ignore_button, QDialogButtonBox.ActionRole)

        def ignore_selected() -> None:
            selected_items = list_widget.selectedItems()
            if not selected_items:
                return
            for item in selected_items:
                warning = item.data(Qt.UserRole)
                if isinstance(warning, core.NameTypoWarning):
                    self.ignored_name_typos.add(core.typo_warning_key(warning.employee, warning.similar_employee))
            self._persist_ignored_name_typos()
            dialog.accept()

        ignore_button.clicked.connect(ignore_selected)
        buttons.rejected.connect(dialog.reject)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)
        dialog.exec()

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
            self._set_update_status(f"Update available: {latest_tag or latest_version} (installed: {core.APP_VERSION})")
            installer_asset = core.choose_release_installer_asset(release_info)
            can_auto_download = installer_asset is not None and self.app_settings.auto_download_updates_on_unmetered_wifi and core.is_unmetered_wifi_profile(network_profile or {})
            if can_auto_download:
                self._start_update_download(release_info, manual=manual)
                if manual:
                    QMessageBox.information(
                        self,
                        "Check for Updates",
                        "A newer version is available\n\n"
                        f"Installed: {core.APP_VERSION}\n"
                        f"Latest: {latest_tag or latest_version}\n"
                        f"Published: {published_at or 'Unknown'}\n"
                        f"Assets: {assets_text}\n\n"
                        "This machine is on unmetered Wi-Fi, so the installer is being downloaded automatically.",
                    )
                return
            network_text = core.describe_network_profile(network_profile or {}) if network_profile is not None else "network state unavailable"
            self._set_update_status(f"Update available: {latest_tag or latest_version} (installed: {core.APP_VERSION}; {network_text})")
            if manual:
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
                        "Automatic download is only offered on unmetered Wi-Fi.\n\n"
                        "Download the installer to this PC now?",
                    ) == QMessageBox.Yes:
                        self._start_update_download(release_info, manual=manual)
                    else:
                        QMessageBox.information(self, "Check for Updates", "You can install later from the release page:\n\n" + html_url)
                else:
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
            core.download_release_asset(download_url, destination)
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
        if manual:
            QMessageBox.critical(self, "Update Download", f"Could not download the update.\n\n{message}")

    def _handle_update_download_success(self, release_info: dict[str, object], destination: Path, checksum_verified: bool, manual: bool) -> None:
        self._update_download_in_progress = False
        self._downloaded_update_path = destination
        latest_tag = str(release_info.get("tag_name", "")).strip() or str(release_info.get("version", "")).strip() or destination.name
        verification_text = " and verified" if checksum_verified else ""
        self._set_update_status(f"Downloaded update {latest_tag}{verification_text}: {destination.name}")
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
        creationflags = 0
        startupinfo = None
        if os.name == "nt":
            creationflags = (
                getattr(subprocess, "DETACHED_PROCESS", 0)
                | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
                | getattr(subprocess, "CREATE_NO_WINDOW", 0)
            )
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = getattr(subprocess, "SW_HIDE", 0)
        try:
            subprocess.Popen(
                ["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", command],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                stdin=subprocess.DEVNULL,
                close_fds=False,
                creationflags=creationflags,
                startupinfo=startupinfo,
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
    app = QApplication.instance() or QApplication(sys.argv)
    if os.name == "nt" and getattr(sys, "frozen", False):
        core._windows_set_explicit_app_user_model_id()
    window = DssQtMainWindow(initial_source=initial_source)
    window.show()
    return app.exec()
