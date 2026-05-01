"""Shared workbook / workspace helpers for DSS Tools tests."""

from __future__ import annotations

import uuid
from contextlib import contextmanager
from pathlib import Path
import unittest

from openpyxl import Workbook


class DssToolsFixtures(unittest.TestCase):
    """Base with temp paths and minimal DSS workbooks (no test methods)."""

    @contextmanager
    def workspace_files(self, stem: str):
        source = Path.cwd() / f"{stem}_{uuid.uuid4().hex}.xlsx"
        try:
            yield source
        finally:
            if source.exists():
                source.unlink()

    @contextmanager
    def workspace_json(self, stem: str):
        path = Path.cwd() / f"{stem}_{uuid.uuid4().hex}.json"
        try:
            yield path
        finally:
            if path.exists():
                path.unlink()

    @contextmanager
    def workspace_dir(self, stem: str):
        path = Path.cwd() / f"{stem}_{uuid.uuid4().hex}"
        path.mkdir(parents=True, exist_ok=True)
        try:
            yield path
        finally:
            if path.exists():
                for child in sorted(path.rglob("*"), reverse=True):
                    if child.is_file():
                        child.unlink()
                    elif child.is_dir():
                        child.rmdir()
                path.rmdir()

    def build_source_workbook(self, path: Path) -> None:
        wb = Workbook()
        ws = wb.active
        ws.title = "2026-04-07 R1"

        ws["AA25"] = 12345
        ws["T25"] = "Alice Smith"
        ws["AC25"] = 8
        ws["AD25"] = 1
        ws["AE25"] = 1
        ws["AC26"] = 2
        ws["AC27"] = 1

        ws["AT28"] = "Bob Jones"
        ws["AX28"] = 4
        ws["AY28"] = 4
        ws["AX29"] = 2
        ws["AZ30"] = 1

        ws2 = wb.create_sheet("2026-04-08 ")
        ws2["T31"] = "Alice Smith"
        ws2["AC31"] = 10
        ws2["AC32"] = 3

        ws3 = wb.create_sheet("Notes")
        ws3["A1"] = "ignore me"

        wb.save(path)

    def build_second_source_workbook(self, path: Path) -> None:
        wb = Workbook()
        ws = wb.active
        ws.title = "2026-04-09 Phase 2"

        ws["T25"] = "Alice Smith"
        ws["AC25"] = 5
        ws["AC26"] = 1

        ws["AT28"] = "Charlie West"
        ws["AX28"] = 7

        wb.save(path)

    def build_revision_source_workbook(self, path: Path) -> None:
        wb = Workbook()
        ws = wb.active
        ws.title = "2026-04-07"
        ws["T25"] = "Alice Smith"
        ws["AC25"] = 8

        rev_ws = wb.create_sheet("2026-04-07 R1")
        rev_ws["T25"] = "Alice Smith"
        rev_ws["AC25"] = 10

        base2_ws = wb.create_sheet("2026-04-08")
        base2_ws["T25"] = "Bob Jones"
        base2_ws["AC25"] = 6

        rev2_ws = wb.create_sheet("2026-04-08 rev 2")
        rev2_ws["T25"] = "Bob Jones"
        rev2_ws["AC25"] = 12

        wb.save(path)

    def build_multiweek_source_workbook(self, path: Path) -> None:
        wb = Workbook()
        sheets = [
            ("2026-04-08", "Alice Smith", 8),
            ("2026-04-15", "Alice Smith", 9),
            ("2026-04-20", "Alice Smith", 10),
            ("2026-04-22", "Alice Smith", 11),
        ]
        for index, (sheet_name, employee, st_hours) in enumerate(sheets):
            ws = wb.active if index == 0 else wb.create_sheet()
            ws.title = sheet_name
            ws["T25"] = employee
            ws["AC25"] = st_hours

        wb.save(path)
