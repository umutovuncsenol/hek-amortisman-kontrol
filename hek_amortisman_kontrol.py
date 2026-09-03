from __future__ import annotations

import os
import re
import subprocess
import sys
import threading
import tkinter as tk
from collections import defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    Workbook = None
    load_workbook = None


APP_TITLE = "HEK / Amortisman Eklenti Kontrolü"
SUPPORTED_TYPES = [("Excel dosyaları", "*.xlsx *.xlsm"), ("Tüm dosyalar", "*.*")]

HEK_PROCESS_COL = 37  # AK
HEK_INVENTORY_COL = 8  # H
HEK_SUBNUMBER_COL = 9  # I
HEK_RATIO_COL = 16  # P
AMORT_INVENTORY_COL = 1  # A
AMORT_SUBNUMBER_COL = 2  # B

YELLOW = "#F9D84A"
PALE_YELLOW = "#FFF8D6"
BLUE = "#2867A8"
PALE_BLUE = "#EAF3FB"
DARK = "#252A31"
GREEN = "#DDF2E3"
RED = "#FCE1E1"
ORANGE = "#FDE9C8"
WHITE = "#FFFFFF"


def xl_color(color: str) -> str:
    return color.lstrip("#")


def resource_path(relative: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / relative


def clean_identifier(value) -> str | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    return text or None


def identifier_key(value) -> str | None:
    text = clean_identifier(value)
    return text.casefold() if text is not None else None


def subnumber_key(value) -> str | None:
    text = clean_identifier(value)
    if text is None:
        return None
    compact = text.replace(" ", "")
    if re.fullmatch(r"[+-]?\d+(?:\.0+)?", compact):
        try:
            return str(int(Decimal(compact)))
        except (InvalidOperation, ValueError):
            pass
    return compact.casefold()


def natural_sort_key(value: str):
    return [int(part) if part.isdigit() else part.casefold() for part in re.split(r"(\d+)", value)]


def parse_process_input(text: str):
    values = []
    seen = set()
    duplicates = 0
    ignored_headers = {"process id", "processid", "süreç no", "süreç numarası", "surec no", "surec numarasi"}

    for raw_line in text.replace("\r", "\n").split("\n"):
        first_cell = raw_line.split("\t", 1)[0].strip()
        if not first_cell:
            continue
        if " ".join(first_cell.split()).casefold() in ignored_headers:
            continue
        key = identifier_key(first_cell)
        if key in seen:
            duplicates += 1
            continue
        seen.add(key)
        values.append((key, first_cell))

    return values, duplicates


def ratio_info(value):
    if value is None or str(value).strip() == "":
        return "(boş)", False

    if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
        number = Decimal(str(value))
        percent = number * 100 if abs(number) <= 1 else number
    else:
        text = str(value).strip().replace("\u00a0", "").replace(" ", "")
        text = text.replace("%", "").replace(",", ".")
        try:
            percent = Decimal(text)
        except InvalidOperation:
            return str(value).strip(), False

    display = f"%{percent.normalize():f}"
    return display, percent == Decimal("100")


def trim_trailing_empty(values):
    result = list(values)
    while result and result[-1] is None:
        result.pop()
    return result


def display_header(value, index: int) -> str:
    text = clean_identifier(value)
    return text or f"Sütun {get_column_letter(index + 1)}"


def ordered_unique(values):
    seen = set()
    result = []
    for value in values:
        if value not in seen:
            seen.add(value)
            result.append(value)
    return result


def unique_join(values, separator=" | "):
    cleaned = [str(value) for value in values if value not in (None, "")]
    return separator.join(ordered_unique(cleaned))


def read_header(path: Path, sheet_name: str, header_row: int):
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        worksheet = workbook[sheet_name]
        row = next(worksheet.iter_rows(min_row=header_row, max_row=header_row, values_only=True), ())
        return trim_trailing_empty(row)
    finally:
        workbook.close()


def collect_hek(path: Path, sheet_name: str, header_row: int, requested_processes):
    requested_keys = {key for key, _display in requested_processes}
    requested_display = {key: display for key, display in requested_processes}
    found_processes = set()
    entries = []
    invalid_rows = []

    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        worksheet = workbook[sheet_name]
        for row_number, row in enumerate(
            worksheet.iter_rows(min_row=header_row + 1, values_only=True), start=header_row + 1
        ):
            process_value = row[HEK_PROCESS_COL - 1] if len(row) >= HEK_PROCESS_COL else None
            process_key = identifier_key(process_value)
            if process_key not in requested_keys:
                continue

            found_processes.add(process_key)
            inventory_value = row[HEK_INVENTORY_COL - 1] if len(row) >= HEK_INVENTORY_COL else None
            subnumber_value = row[HEK_SUBNUMBER_COL - 1] if len(row) >= HEK_SUBNUMBER_COL else None
            ratio_value = row[HEK_RATIO_COL - 1] if len(row) >= HEK_RATIO_COL else None
            inventory_key = identifier_key(inventory_value)
            sub_key = subnumber_key(subnumber_value)
            ratio_display, ratio_full = ratio_info(ratio_value)

            if inventory_key is None or sub_key is None:
                invalid_rows.append({
                    "process": requested_display[process_key],
                    "row": row_number,
                    "inventory": clean_identifier(inventory_value) or "(boş)",
                    "subnumber": clean_identifier(subnumber_value) or "(boş)",
                    "ratio": ratio_display,
                    "reason": "HEK satırında envanter veya eklenti numarası boş",
                })
                continue

            entries.append({
                "process_key": process_key,
                "process": requested_display[process_key],
                "inventory_key": inventory_key,
                "inventory": clean_identifier(inventory_value),
                "sub_key": sub_key,
                "subnumber": clean_identifier(subnumber_value),
                "ratio": ratio_display,
                "ratio_full": ratio_full,
                "row": row_number,
            })
    finally:
        workbook.close()

    missing_processes = [display for key, display in requested_processes if key not in found_processes]
    return entries, invalid_rows, missing_processes


def collect_amortisation(path: Path, sheet_name: str, header_row: int, inventory_keys):
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        worksheet = workbook[sheet_name]
        header_values = next(
            worksheet.iter_rows(min_row=header_row, max_row=header_row, values_only=True), ()
        )
        headers = trim_trailing_empty(header_values)
        source_column_count = max(len(headers), 2)
        headers.extend([None] * (source_column_count - len(headers)))

        rows_by_inventory = defaultdict(list)
        for row_number, row in enumerate(
            worksheet.iter_rows(min_row=header_row + 1, values_only=True), start=header_row + 1
        ):
            inventory_value = row[AMORT_INVENTORY_COL - 1] if row else None
            inventory_key_value = identifier_key(inventory_value)
            if inventory_key_value not in inventory_keys:
                continue
            padded = list(row[:source_column_count])
            padded.extend([None] * (source_column_count - len(padded)))
            rows_by_inventory[inventory_key_value].append({
                "values": padded,
                "inventory": clean_identifier(inventory_value),
                "subnumber": clean_identifier(row[AMORT_SUBNUMBER_COL - 1]) if len(row) >= 2 else None,
                "sub_key": subnumber_key(row[AMORT_SUBNUMBER_COL - 1]) if len(row) >= 2 else None,
                "row": row_number,
            })

        return headers, rows_by_inventory
    finally:
        workbook.close()


def build_analysis(processes, hek_entries, invalid_hek_rows, missing_processes, headers, amort_rows):
    entries_by_inventory = defaultdict(list)
    request_by_pair = defaultdict(list)
    inventory_display = {}

    for entry in hek_entries:
        entries_by_inventory[entry["inventory_key"]].append(entry)
        request_by_pair[(entry["inventory_key"], entry["sub_key"])].append(entry)
        inventory_display.setdefault(entry["inventory_key"], entry["inventory"])

    summary_rows = []
    detail_groups = []
    unmatched_rows = list(invalid_hek_rows)

    sorted_inventories = sorted(entries_by_inventory, key=lambda key: natural_sort_key(inventory_display[key]))
    for inventory_key_value in sorted_inventories:
        hek_for_inventory = entries_by_inventory[inventory_key_value]
        process_names = ordered_unique(entry["process"] for entry in hek_for_inventory)
        rows = list(amort_rows.get(inventory_key_value, []))
        rows.sort(key=lambda item: natural_sort_key(item["sub_key"] or ""))
        available_pairs = {(inventory_key_value, row["sub_key"]) for row in rows}

        for entry in hek_for_inventory:
            if (inventory_key_value, entry["sub_key"]) not in available_pairs:
                unmatched_rows.append({
                    "process": entry["process"],
                    "row": entry["row"],
                    "inventory": entry["inventory"],
                    "subnumber": entry["subnumber"],
                    "ratio": entry["ratio"],
                    "reason": "HEK kaydı Amortisman dosyasında bulunamadı",
                })

        addon_rows = [row for row in rows if row["sub_key"] not in (None, "0")]
        requested_addons = [
            row for row in addon_rows if request_by_pair.get((inventory_key_value, row["sub_key"]))
        ]
        missing_addons = [
            row for row in addon_rows if not request_by_pair.get((inventory_key_value, row["sub_key"]))
        ]
        completion = len(requested_addons) / len(addon_rows) if addon_rows else 1.0
        partial_entries = [entry for entry in hek_for_inventory if not entry["ratio_full"]]

        problems = []
        if not rows:
            problems.append("Envanter Amortisman dosyasında bulunamadı")
        if missing_addons:
            problems.append(f"{len(missing_addons)} kalan eklenti var")
        if partial_entries:
            problems.append("P oranı %100 değil")
        status = "Kontrol gerekli" if problems else "Tamam"

        summary_rows.append({
            "processes": "\n".join(process_names),
            "inventory": inventory_display[inventory_key_value],
            "addon_total": len(addon_rows),
            "requested_count": len(requested_addons),
            "requested_list": unique_join(row["subnumber"] for row in requested_addons) or "—",
            "missing_count": len(missing_addons),
            "missing_list": unique_join(row["subnumber"] for row in missing_addons) or "—",
            "completion": completion,
            "p_check": unique_join(
                f'{entry["process"]}: {entry["ratio"]}' for entry in partial_entries
            ) or "%100",
            "status": status,
            "explanation": "; ".join(problems) or "Bütün eklentiler hurdalanmak istenmiş",
        })

        detail_rows = []
        for row in rows:
            requests = request_by_pair.get((inventory_key_value, row["sub_key"]), [])
            is_main = row["sub_key"] == "0"
            if requests:
                requesting_processes = ordered_unique(entry["process"] for entry in requests)
                ratios = ordered_unique(entry["ratio"] for entry in requests)
                partial = any(not entry["ratio_full"] for entry in requests)
                if is_main:
                    status_text = "Ana ürün - hurdalanmak isteniyor"
                else:
                    status_text = "Hurdalanmak istenen eklenti"
                explanation = "P oranı %100 değil" if partial else "HEK Formunda mevcut"
                kind = "partial" if partial else "requested"
            else:
                requesting_processes = []
                ratios = []
                if is_main:
                    status_text = "Ana ürün HEK Formunda yok"
                    explanation = "Ana ürün satırı hurdalama talebinde bulunamadı"
                    kind = "warning"
                else:
                    status_text = "Kalan eklenti"
                    explanation = "Hurdalanmak istenmemiş / eksik eklenti"
                    kind = "missing"

            detail_rows.append({
                **row,
                "request_status": status_text,
                "request_processes": "\n".join(requesting_processes) or "—",
                "ratios": unique_join(ratios) or "—",
                "explanation": explanation,
                "kind": kind,
            })

        detail_groups.append({
            "processes": "\n".join(process_names),
            "inventory": inventory_display[inventory_key_value],
            "rows": detail_rows,
        })

    for process in missing_processes:
        unmatched_rows.append({
            "process": process,
            "row": "—",
            "inventory": "—",
            "subnumber": "—",
            "ratio": "—",
            "reason": "Süreç numarası HEK Formunda bulunamadı",
        })

    return {
        "headers": headers,
        "summary": summary_rows,
        "detail_groups": detail_groups,
        "unmatched": unmatched_rows,
        "process_count": len(processes),
    }


def style_header(row_cells, fill_color=BLUE):
    for cell in row_cells:
        cell.fill = PatternFill("solid", fgColor=xl_color(fill_color))
        cell.font = Font(color=xl_color(WHITE), bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def write_result_workbook(result, output_path: Path):
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Kontrol Özeti"
    detail = workbook.create_sheet("Filtrelenmiş Amortisman")
    unmatched = workbook.create_sheet("Eşleşmeyen Kayıtlar")

    thin_gray = Side(style="thin", color="D4DAE1")
    section_border = Border(bottom=Side(style="medium", color=xl_color(BLUE)))

    summary_headers = [
        "Süreç Numaraları", "Envanter Numarası", "Toplam Eklenti",
        "Hurdalanacak Eklenti", "Hurdalanacak Eklenti No", "Kalan Eklenti",
        "Kalan Eklenti No", "Eklenti Tamamlanma", "P Kontrolü", "Durum", "Açıklama",
    ]
    summary.append(summary_headers)
    style_header(summary[1])
    summary.row_dimensions[1].height = 34

    for item in result["summary"]:
        summary.append([
            item["processes"], item["inventory"], item["addon_total"],
            item["requested_count"], item["requested_list"], item["missing_count"],
            item["missing_list"], item["completion"], item["p_check"], item["status"],
            item["explanation"],
        ])
        row = summary.max_row
        summary.cell(row, 8).number_format = "0%"
        fill = GREEN if item["status"] == "Tamam" else ORANGE
        if item["missing_count"]:
            fill = RED
        for cell in summary[row]:
            cell.fill = PatternFill("solid", fgColor=xl_color(fill))
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = Border(bottom=thin_gray)

    summary.freeze_panes = "A2"
    summary.auto_filter.ref = summary.dimensions
    summary.sheet_view.showGridLines = False
    summary_widths = [25, 19, 14, 18, 25, 14, 25, 18, 28, 18, 42]
    for index, width in enumerate(summary_widths, start=1):
        summary.column_dimensions[get_column_letter(index)].width = width

    source_headers = [display_header(value, index) for index, value in enumerate(result["headers"])]
    detail_headers = ["Süreç Numaraları"] + source_headers + [
        "HEK Talep Durumu", "Talep Eden Süreç(ler)", "txtKismiCikisOrani", "Kontrol Açıklaması"
    ]
    detail.append(detail_headers)
    style_header(detail[1])
    detail.row_dimensions[1].height = 38
    detail.freeze_panes = "B2"
    detail.sheet_view.showGridLines = False

    for group_number, group in enumerate(result["detail_groups"]):
        start_row = detail.max_row + 1
        for item in group["rows"]:
            detail.append([
                group["processes"], *item["values"], item["request_status"],
                item["request_processes"], item["ratios"], item["explanation"],
            ])
            current_row = detail.max_row
            if item["kind"] == "missing":
                fill_color = RED
            elif item["kind"] in {"partial", "warning"}:
                fill_color = ORANGE
            else:
                fill_color = GREEN if group_number % 2 == 0 else PALE_BLUE
            for cell in detail[current_row]:
                cell.fill = PatternFill("solid", fgColor=xl_color(fill_color))
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            detail.cell(current_row, 1).border = section_border

        end_row = detail.max_row
        if end_row >= start_row:
            if end_row > start_row:
                detail.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
            process_cell = detail.cell(start_row, 1)
            process_cell.value = group["processes"]
            process_cell.font = Font(bold=True, color=xl_color(DARK))
            process_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            process_cell.fill = PatternFill("solid", fgColor=xl_color(YELLOW))
            process_cell.border = Border(
                left=Side(style="medium", color=xl_color(BLUE)),
                right=Side(style="medium", color=xl_color(BLUE)),
                top=Side(style="medium", color=xl_color(BLUE)),
                bottom=Side(style="medium", color=xl_color(BLUE)),
            )

    if detail.max_row > 1:
        detail.auto_filter.ref = f"B1:{get_column_letter(detail.max_column)}{detail.max_row}"
    detail.column_dimensions["A"].width = 25
    for index, header in enumerate(source_headers, start=2):
        folded = header.casefold()
        if any(word in folded for word in ("description", "tanım", "tanim", "açıklama", "aciklama")):
            width = 42
        elif any(word in folded for word in ("date", "tarih")):
            width = 14
        else:
            width = min(max(len(header) + 2, 12), 24)
        detail.column_dimensions[get_column_letter(index)].width = width
    for index, width in enumerate([28, 24, 22, 42], start=2 + len(source_headers)):
        detail.column_dimensions[get_column_letter(index)].width = width

    unmatched_headers = [
        "Süreç Numarası", "HEK Satırı", "Envanter Numarası", "Eklenti Numarası",
        "txtKismiCikisOrani", "Açıklama",
    ]
    unmatched.append(unmatched_headers)
    style_header(unmatched[1])
    for item in result["unmatched"]:
        unmatched.append([
            item["process"], item["row"], item["inventory"], item["subnumber"],
            item["ratio"], item["reason"],
        ])
        for cell in unmatched[unmatched.max_row]:
            cell.fill = PatternFill("solid", fgColor=xl_color(RED))
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    unmatched.freeze_panes = "A2"
    unmatched.auto_filter.ref = unmatched.dimensions
    unmatched.sheet_view.showGridLines = False
    for index, width in enumerate([24, 12, 20, 18, 22, 48], start=1):
        unmatched.column_dimensions[get_column_letter(index)].width = width

    for sheet in (summary, detail, unmatched):
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 0
        sheet.sheet_properties.outlinePr.summaryBelow = True

    workbook.save(output_path)


def run_analysis(process_text, hek_config, amort_config, output_path: Path):
    processes, duplicates = parse_process_input(process_text)
    if not processes:
        raise ValueError("En az bir süreç numarası yapıştırın.")

    hek_entries, invalid_rows, missing_processes = collect_hek(
        hek_config["path"], hek_config["sheet"], hek_config["header_row"], processes
    )
    inventory_keys = {entry["inventory_key"] for entry in hek_entries}
    headers, amort_rows = collect_amortisation(
        amort_config["path"], amort_config["sheet"], amort_config["header_row"], inventory_keys
    )
    result = build_analysis(
        processes, hek_entries, invalid_rows, missing_processes, headers, amort_rows
    )
    result["duplicates"] = duplicates
    result["hek_entry_count"] = len(hek_entries)
    result["inventory_count"] = len(inventory_keys)
    result["missing_addon_count"] = sum(item["missing_count"] for item in result["summary"])
    result["partial_inventory_count"] = sum(
        1 for item in result["summary"] if item["p_check"] != "%100"
    )
    write_result_workbook(result, output_path)
    return result


class ExcelCard(ttk.LabelFrame):
    def __init__(self, master, title, kind):
        super().__init__(master, text=title, padding=12, style="Card.TLabelframe")
        self.kind = kind
        self.path: Path | None = None

        self.columnconfigure(1, weight=1)
        ttk.Label(self, text="Dosya", style="Card.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.file_var = tk.StringVar(value="Dosya seçilmedi")
        ttk.Entry(self, textvariable=self.file_var, state="readonly").grid(row=0, column=1, sticky="ew")
        ttk.Button(self, text="Dosya Seç", command=self.choose_file, style="Blue.TButton").grid(
            row=0, column=2, padx=(8, 0)
        )

        ttk.Label(self, text="Sayfa", style="Card.TLabel").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.sheet_var = tk.StringVar()
        self.sheet_box = ttk.Combobox(self, textvariable=self.sheet_var, state="readonly")
        self.sheet_box.grid(row=1, column=1, sticky="ew", pady=(8, 0))
        self.sheet_box.bind("<<ComboboxSelected>>", lambda _event: self.validate_columns())

        ttk.Label(self, text="Başlık satırı", style="Card.TLabel").grid(
            row=1, column=2, sticky="e", padx=(8, 76), pady=(8, 0)
        )
        self.header_var = tk.IntVar(value=1)
        ttk.Spinbox(self, from_=1, to=1000, width=6, textvariable=self.header_var,
                    command=self.validate_columns).grid(row=1, column=2, sticky="e", pady=(8, 0))

        self.status_var = tk.StringVar(value="Dosya bekleniyor")
        ttk.Label(self, textvariable=self.status_var, style="Hint.TLabel", wraplength=780).grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(9, 0)
        )

    def choose_file(self):
        filename = filedialog.askopenfilename(title=f"{self['text']} seç", filetypes=SUPPORTED_TYPES)
        if not filename:
            return
        path = Path(filename)
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            messagebox.showerror("Desteklenmeyen dosya", "Lütfen .xlsx veya .xlsm dosyası seçin.")
            return
        try:
            workbook = load_workbook(path, read_only=True, data_only=True)
            sheet_names = workbook.sheetnames
            workbook.close()
        except Exception as exc:
            messagebox.showerror("Dosya açılamadı", f"Excel dosyası okunamadı:\n{exc}")
            return
        self.path = path
        self.file_var.set(str(path))
        self.sheet_box["values"] = sheet_names
        self.sheet_var.set(sheet_names[0])
        self.validate_columns()

    def validate_columns(self):
        if not self.path or not self.sheet_var.get():
            return False
        try:
            header_row = int(self.header_var.get())
            headers = read_header(self.path, self.sheet_var.get(), header_row)
        except Exception as exc:
            self.status_var.set(f"⚠ Başlıklar okunamadı: {exc}")
            return False

        if self.kind == "hek":
            required = [
                (HEK_PROCESS_COL, "AK / Process ID"),
                (HEK_INVENTORY_COL, "H / Envanter"),
                (HEK_SUBNUMBER_COL, "I / Eklenti"),
                (HEK_RATIO_COL, "P / txtKismiCikisOrani"),
            ]
        else:
            required = [
                (AMORT_INVENTORY_COL, "A / Immobilisation"),
                (AMORT_SUBNUMBER_COL, "B / Numero subsidiaire"),
            ]

        missing = [label for index, label in required if len(headers) < index]
        if missing:
            self.status_var.set("⚠ Eksik sütun: " + ", ".join(missing))
            return False

        detected = []
        for index, label in required:
            detected.append(f"{label}: {display_header(headers[index - 1], index - 1)}")
        self.status_var.set("✓ " + "   |   ".join(detected))
        return True

    def config(self):
        if not self.path:
            raise ValueError(f"{self['text']} seçilmedi.")
        if not self.validate_columns():
            raise ValueError(f"{self['text']} sütunları doğrulanamadı.")
        return {
            "path": self.path,
            "sheet": self.sheet_var.get(),
            "header_row": int(self.header_var.get()),
        }


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1120x820")
        self.minsize(900, 720)
        self.configure(bg=PALE_YELLOW)
        self.last_output: Path | None = None
        self._configure_styles()

        shell = ttk.Frame(self, padding=18, style="Shell.TFrame")
        shell.pack(fill="both", expand=True)
        shell.columnconfigure(0, weight=1)
        shell.rowconfigure(4, weight=1)

        header = ttk.Frame(shell, style="Header.TFrame", padding=14)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        header.columnconfigure(1, weight=1)
        self._load_logo(header)
        ttk.Label(header, text="HEK / Amortisman Eklenti Kontrolü", style="Title.TLabel").grid(
            row=0, column=1, sticky="w"
        )
        ttk.Label(
            header,
            text="Süreçleri bulur, envanter eklentilerini kontrol eder ve düzenli Excel raporu üretir.",
            style="Subtitle.TLabel",
        ).grid(row=1, column=1, sticky="w", pady=(4, 0))

        process_card = ttk.LabelFrame(shell, text="1  Süreç Numaraları", padding=12, style="Card.TLabelframe")
        process_card.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        process_card.columnconfigure(0, weight=1)
        self.process_text = scrolledtext.ScrolledText(
            process_card, height=6, font=("TkDefaultFont", 11), wrap="none",
            bg=WHITE, fg=DARK, insertbackground=DARK, relief="flat", borderwidth=1,
        )
        self.process_text.grid(row=0, column=0, columnspan=3, sticky="ew")
        self.process_text.bind("<KeyRelease>", lambda _event: self.update_process_count())
        ttk.Button(process_card, text="Panodan Yapıştır", command=self.paste_processes,
                   style="Yellow.TButton").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Button(process_card, text="Temizle", command=self.clear_processes).grid(
            row=1, column=1, sticky="w", padx=8, pady=(8, 0)
        )
        self.process_count_var = tk.StringVar(value="Süreç numaralarını Excel'den kopyalayıp buraya yapıştırın")
        ttk.Label(process_card, textvariable=self.process_count_var, style="Hint.TLabel").grid(
            row=1, column=2, sticky="e", pady=(8, 0)
        )

        cards = ttk.Frame(shell, style="Shell.TFrame")
        cards.grid(row=2, column=0, sticky="ew")
        cards.columnconfigure(0, weight=1)
        cards.columnconfigure(1, weight=1)
        self.hek_card = ExcelCard(cards, "2  HEK Formu", "hek")
        self.amort_card = ExcelCard(cards, "3  Amortisman Dosyası", "amort")
        self.hek_card.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        self.amort_card.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        action = ttk.Frame(shell, style="Shell.TFrame")
        action.grid(row=3, column=0, sticky="ew", pady=12)
        self.run_button = ttk.Button(action, text="ANALİZİ BAŞLAT", command=self.start_analysis,
                                     style="Primary.TButton")
        self.run_button.pack(side="left")
        self.open_button = ttk.Button(action, text="Sonuç Excel'ini Aç", command=self.open_output,
                                      state="disabled")
        self.open_button.pack(side="left", padx=8)
        self.folder_button = ttk.Button(action, text="Klasörde Göster", command=self.show_output_folder,
                                        state="disabled")
        self.folder_button.pack(side="left")

        self.progress = ttk.Progressbar(action, mode="indeterminate", length=210)
        self.progress.pack(side="right")

        result_card = ttk.LabelFrame(shell, text="Sonuç", padding=16, style="Card.TLabelframe")
        result_card.grid(row=4, column=0, sticky="nsew")
        result_card.columnconfigure(0, weight=1)
        result_card.rowconfigure(1, weight=1)
        self.status_var = tk.StringVar(value="Hazır")
        ttk.Label(result_card, textvariable=self.status_var, style="ResultTitle.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        self.result_var = tk.StringVar(value="Üç girişi tamamlayıp analizi başlatın.")
        ttk.Label(result_card, textvariable=self.result_var, style="Result.TLabel",
                  justify="left", wraplength=980).grid(row=1, column=0, sticky="nw", pady=(12, 0))

    def _configure_styles(self):
        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("Shell.TFrame", background=PALE_YELLOW)
        style.configure("Header.TFrame", background=BLUE)
        style.configure("Title.TLabel", background=BLUE, foreground=WHITE, font=("TkDefaultFont", 19, "bold"))
        style.configure("Subtitle.TLabel", background=BLUE, foreground="#EAF3FB", font=("TkDefaultFont", 10))
        style.configure("Card.TLabelframe", background=WHITE, bordercolor=BLUE, relief="solid")
        style.configure("Card.TLabelframe.Label", background=YELLOW, foreground=DARK,
                        font=("TkDefaultFont", 11, "bold"))
        style.configure("Card.TLabel", background=WHITE, foreground=DARK)
        style.configure("Hint.TLabel", background=WHITE, foreground="#4E5965")
        style.configure("ResultTitle.TLabel", background=WHITE, foreground=BLUE,
                        font=("TkDefaultFont", 15, "bold"))
        style.configure("Result.TLabel", background=WHITE, foreground=DARK, font=("TkDefaultFont", 11))
        style.configure("Primary.TButton", background=BLUE, foreground=WHITE,
                        font=("TkDefaultFont", 12, "bold"), padding=(18, 10))
        style.map("Primary.TButton", background=[("active", "#1F527F")])
        style.configure("Yellow.TButton", background=YELLOW, foreground=DARK, padding=(12, 6))
        style.map("Yellow.TButton", background=[("active", "#E8C62E")])
        style.configure("Blue.TButton", background=BLUE, foreground=WHITE)
        style.map("Blue.TButton", background=[("active", "#1F527F")])

    def _load_logo(self, parent):
        try:
            image = tk.PhotoImage(file=str(resource_path("assets/renault-logo.gif")))
            self.icon_image = image
            self.logo_image = image.subsample(4, 4)
            logo_frame = tk.Frame(parent, bg=WHITE, padx=5, pady=4)
            logo_frame.grid(row=0, column=0, rowspan=2, padx=(0, 16))
            tk.Label(logo_frame, image=self.logo_image, bg=WHITE, borderwidth=0).pack()
            self.iconphoto(True, image)
        except Exception:
            ttk.Label(parent, text="◉", style="Title.TLabel", font=("TkDefaultFont", 30, "bold")).grid(
                row=0, column=0, rowspan=2, padx=(0, 16)
            )

    def paste_processes(self):
        try:
            value = self.clipboard_get()
        except tk.TclError:
            messagebox.showinfo("Pano boş", "Panoda yapıştırılabilecek metin bulunamadı.")
            return
        self.process_text.delete("1.0", "end")
        self.process_text.insert("1.0", value)
        self.update_process_count()

    def clear_processes(self):
        self.process_text.delete("1.0", "end")
        self.update_process_count()

    def update_process_count(self):
        processes, duplicates = parse_process_input(self.process_text.get("1.0", "end"))
        if not processes:
            self.process_count_var.set("Süreç numaralarını Excel'den kopyalayıp buraya yapıştırın")
        else:
            suffix = f" • {duplicates} tekrar çıkarıldı" if duplicates else ""
            self.process_count_var.set(f"{len(processes)} benzersiz süreç{suffix}")

    def start_analysis(self):
        try:
            process_text = self.process_text.get("1.0", "end")
            processes, _duplicates = parse_process_input(process_text)
            if not processes:
                raise ValueError("Süreç numaralarını Excel'den kopyalayıp yapıştırın.")
            hek_config = self.hek_card.config()
            amort_config = self.amort_card.config()
        except Exception as exc:
            messagebox.showerror("Eksik bilgi", str(exc))
            return

        filename = filedialog.asksaveasfilename(
            title="Analiz Excel dosyasını kaydet",
            defaultextension=".xlsx",
            initialfile="HEK_Amortisman_Eklenti_Kontrolu.xlsx",
            filetypes=[("Excel dosyası", "*.xlsx")],
        )
        if not filename:
            return

        self.run_button.configure(state="disabled")
        self.open_button.configure(state="disabled")
        self.folder_button.configure(state="disabled")
        self.progress.start(10)
        self.status_var.set("Analiz yapılıyor…")
        self.result_var.set("HEK Formu ve Amortisman dosyası birer kez taranıyor. Lütfen bekleyin.")

        thread = threading.Thread(
            target=self._analysis_worker,
            args=(process_text, hek_config, amort_config, Path(filename)),
            daemon=True,
        )
        thread.start()

    def _analysis_worker(self, process_text, hek_config, amort_config, output_path):
        try:
            result = run_analysis(process_text, hek_config, amort_config, output_path)
        except Exception as exc:
            self.after(0, lambda: self._analysis_failed(exc))
            return
        self.after(0, lambda: self._analysis_finished(result, output_path))

    def _analysis_failed(self, exc):
        self.progress.stop()
        self.run_button.configure(state="normal")
        self.status_var.set("Analiz tamamlanamadı")
        self.result_var.set(str(exc))
        messagebox.showerror("Analiz tamamlanamadı", str(exc))

    def _analysis_finished(self, result, output_path):
        self.progress.stop()
        self.run_button.configure(state="normal")
        self.open_button.configure(state="normal")
        self.folder_button.configure(state="normal")
        self.last_output = output_path
        matched_processes = result["process_count"] - len([
            row for row in result["unmatched"] if row["reason"] == "Süreç numarası HEK Formunda bulunamadı"
        ])
        self.status_var.set("Analiz tamamlandı")
        self.result_var.set(
            f"✓ {matched_processes}/{result['process_count']} süreç HEK Formunda bulundu\n"
            f"✓ {result['inventory_count']} envanter numarası analiz edildi\n"
            f"⚠ {result['missing_addon_count']} kalan eklenti bulundu\n"
            f"⚠ {result['partial_inventory_count']} envanter grubunda P oranı %100 değil\n"
            f"⚠ {len(result['unmatched'])} eşleşmeyen veya eksik kayıt var\n\n"
            f"Excel kaydedildi: {output_path}"
        )
        messagebox.showinfo("Tamamlandı", "Analiz Excel dosyası başarıyla oluşturuldu.")

    def open_output(self):
        if self.last_output:
            self._open_path(self.last_output)

    def show_output_folder(self):
        if not self.last_output:
            return
        if sys.platform == "win32":
            subprocess.Popen(["explorer", "/select,", str(self.last_output)])
        elif sys.platform == "darwin":
            subprocess.Popen(["open", "-R", str(self.last_output)])
        else:
            subprocess.Popen(["xdg-open", str(self.last_output.parent)])

    @staticmethod
    def _open_path(path: Path):
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])


if __name__ == "__main__":
    if load_workbook is None:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Eksik bileşen",
            "Excel okuyucu bulunamadı. Önce openpyxl paketini kurun.",
        )
        raise SystemExit(1)
    App().mainloop()
