from __future__ import annotations

"""FileHound - safe Windows file finder."""

import csv
import difflib
import os
import re
import shutil
import tempfile
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk

import pandas as pd
import piexif
from openpyxl.styles import Alignment
from PIL import Image, ImageTk


APP_NAME = "FileHound"
APP_TITLE = "FileHound"
APP_VERSION = "1.0.0"
APP_AUTHOR = "f-estero"
APP_REPOSITORY = "https://github.com/f-estero/filehound.git"
WINDOW_SIZE = "1180x820"
JPEG_EXTENSIONS = {".jpg", ".jpeg"}
SUPPORTED_IMPORT_EXTENSIONS = {".xlsx", ".xls", ".csv"}
SEARCH_COLUMNS = ("Request", "File", "Path", "Extension", "Size", "Modified", "Match", "Status")
POTENTIALLY_DANGEROUS_EXTENSIONS = {
    ".exe", ".bat", ".cmd", ".vbs", ".vbe", ".js", ".jse", ".wsf",
    ".wsh", ".ps1", ".ps1xml", ".ps2", ".msc", ".msi", ".msp",
    ".com", ".scr", ".hta", ".cpl", ".jar", ".reg", ".pif"
}


def sanitize_spreadsheet_value(value: object) -> object:
    if isinstance(value, str) and value.startswith(("=", "+", "-", "@", "\t", "\r")):
        return f"'{value}"
    return value


@dataclass
class SearchResult:
    requested_name: str
    file_name: str
    path: str
    extension: str
    size_bytes: int
    modified_time: str
    match_type: str
    status: str
    gps_status: str = ""
    latitude: str = ""
    longitude: str = ""


@dataclass
class GpsProcessRecord:
    source_name: str
    output_name: str
    source_path: str
    working_path: str
    inserted_manually: bool
    converted_to_jpeg: bool


def normalize_name(name: str) -> str:
    return str(name).strip().lower()


def normalize_roe_name(name: str) -> str:
    return normalize_name(name).replace("/", "-")


def format_size(size_bytes: int) -> str:
    units = ["B", "KB", "MB", "GB", "TB"]
    size = float(size_bytes)
    for unit in units:
        if size < 1024 or unit == units[-1]:
            return f"{size:.1f} {unit}"
        size /= 1024
    return f"{size_bytes} B"


def safe_stat(path: str) -> tuple[int, str]:
    stats = os.stat(path)
    modified = pd.Timestamp(stats.st_mtime, unit="s").strftime("%Y-%m-%d %H:%M:%S")
    return stats.st_size, modified


def iter_files(source_dir: str) -> list[str]:
    collected: list[str] = []
    for root_dir, _, files in os.walk(source_dir):
        for file_name in files:
            collected.append(os.path.join(root_dir, file_name))
    return collected


def read_tabular_file(file_path: str) -> pd.DataFrame:
    suffix = Path(file_path).suffix.lower()
    if suffix == ".csv":
        encodings = ("utf-8-sig", "utf-8", "cp1252", "latin-1")
        last_error: Exception | None = None
        for encoding in encodings:
            try:
                return pd.read_csv(file_path, encoding=encoding)
            except Exception as exc:
                last_error = exc
        raise ValueError(f"Unable to read CSV file: {last_error}")
    if suffix in {".xlsx", ".xls"}:
        return pd.read_excel(file_path)
    raise ValueError("Unsupported format. Use XLSX, XLS, or CSV.")


def get_unique_destination_path(destination_dir: str, file_name: str) -> str:
    base = Path(file_name).stem
    suffix = Path(file_name).suffix
    candidate = os.path.join(destination_dir, file_name)
    counter = 1
    while os.path.exists(candidate):
        candidate = os.path.join(destination_dir, f"{base}__copy{counter}{suffix}")
        counter += 1
    return candidate


def safe_copy_file(source_path: str, destination_dir: str, output_name: str | None = None) -> str:
    os.makedirs(destination_dir, exist_ok=True)
    raw_name = output_name or os.path.basename(source_path)
    target_name = os.path.basename(raw_name)
    if not target_name:
        target_name = "file"
    destination_path = get_unique_destination_path(destination_dir, target_name)
    dest_abs = os.path.abspath(destination_path)
    dir_abs = os.path.abspath(destination_dir)
    if not dest_abs.startswith(dir_abs):
        raise ValueError("Security error: destination path escapes destination directory.")
    shutil.copy2(source_path, destination_path)
    return destination_path


def find_roe_variants(base_name: str, all_files: list[str]) -> list[str]:
    escaped = re.escape(base_name)
    pattern = re.compile(rf"^{escaped}(?:__\d+)?\.[a-z0-9]+$", re.IGNORECASE)
    matches = [path for path in all_files if pattern.match(os.path.basename(path))]

    def sort_key(path: str) -> tuple[int, str]:
        file_name = os.path.basename(path).lower()
        match = re.search(r"__(\d+)", file_name)
        return (int(match.group(1)) if match else 0, file_name)

    return sorted(matches, key=sort_key)


def gps_tuple_to_decimal(values, ref) -> float:
    if isinstance(ref, bytes):
        ref = ref.decode(errors="ignore")
    degrees = values[0][0] / values[0][1]
    minutes = values[1][0] / values[1][1]
    seconds = values[2][0] / values[2][1]
    decimal = degrees + minutes / 60 + seconds / 3600
    if ref in ("S", "W"):
        decimal = -decimal
    return decimal


def extract_gps_metadata(image_path: str) -> tuple[str, str, str]:
    try:
        with Image.open(image_path) as image:
            exif_data = image.getexif()

        gps_info = exif_data.get(34853)
        if not gps_info:
            return "NO", "", ""

        lat = gps_info.get(2)
        lon = gps_info.get(4)
        if not lat or not lon:
            return "NO", "", ""

        latitude = gps_tuple_to_decimal(lat, gps_info.get(1, "N"))
        longitude = gps_tuple_to_decimal(lon, gps_info.get(3, "E"))
        if -90 <= latitude <= 90 and -180 <= longitude <= 180:
            return "SI", str(latitude), str(longitude)
        return "NO", "", ""
    except Exception:
        return "NO", "", ""


def decimal_to_exif_dms(decimal_value: float) -> list[tuple[int, int]]:
    degrees = int(abs(decimal_value))
    minutes_full = (abs(decimal_value) - degrees) * 60
    minutes = int(minutes_full)
    seconds = (minutes_full - minutes) * 60
    return [
        (degrees, 1),
        (minutes, 1),
        (int(seconds * 1_000_000), 1_000_000),
    ]


def write_gps_metadata(image_path: str, latitude: float, longitude: float) -> None:
    try:
        exif_dict = piexif.load(image_path)
    except Exception:
        exif_dict = {"0th": {}, "Exif": {}, "GPS": {}, "1st": {}, "thumbnail": None}

    exif_dict["GPS"][piexif.GPSIFD.GPSLatitudeRef] = ("N" if latitude >= 0 else "S").encode()
    exif_dict["GPS"][piexif.GPSIFD.GPSLatitude] = decimal_to_exif_dms(latitude)
    exif_dict["GPS"][piexif.GPSIFD.GPSLongitudeRef] = ("E" if longitude >= 0 else "W").encode()
    exif_dict["GPS"][piexif.GPSIFD.GPSLongitude] = decimal_to_exif_dms(longitude)
    piexif.insert(piexif.dump(exif_dict), image_path)


def convert_to_temp_jpeg(source_path: str) -> str:
    with Image.open(source_path) as image:
        rgb_image = image.convert("RGB")
        with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as temp_file:
            rgb_image.save(temp_file.name, "JPEG", quality=95)
            return temp_file.name


class SearchEngine:
    @staticmethod
    def search_by_query(source_dir: str, query: str, mode: str) -> list[SearchResult]:
        query_clean = query.strip()
        if not query_clean:
            return []

        query_normalized = normalize_name(query_clean)
        results: list[SearchResult] = []

        for path in iter_files(source_dir):
            file_name = os.path.basename(path)
            file_normalized = normalize_name(file_name)
            stem_normalized = normalize_name(Path(file_name).stem)
            match_type = ""

            if mode == "Exact":
                if file_normalized == query_normalized or stem_normalized == query_normalized:
                    match_type = "Exact"
            elif mode == "Contains":
                if query_normalized in file_normalized or query_normalized in stem_normalized:
                    match_type = "Partial"
            else:
                score = max(
                    difflib.SequenceMatcher(None, query_normalized, file_normalized).ratio(),
                    difflib.SequenceMatcher(None, query_normalized, stem_normalized).ratio(),
                )
                if score >= 0.72:
                    match_type = f"Similar {score:.2f}"

            if match_type:
                size_bytes, modified_time = safe_stat(path)
                results.append(
                    SearchResult(
                        requested_name=query_clean,
                        file_name=file_name,
                        path=path,
                        extension=Path(file_name).suffix.lower(),
                        size_bytes=size_bytes,
                        modified_time=modified_time,
                        match_type=match_type,
                        status="Found",
                    )
                )

        results.sort(key=lambda item: (item.file_name.lower(), item.path.lower()))
        return results

    @staticmethod
    def search_batch(source_dir: str, requested_names: list[str], mode: str) -> list[SearchResult]:
        aggregated: list[SearchResult] = []
        seen_paths: set[tuple[str, str]] = set()

        for requested_name in requested_names:
            current_results = SearchEngine.search_by_query(source_dir, requested_name, mode)
            if current_results:
                for result in current_results:
                    result_key = (requested_name, result.path)
                    if result_key not in seen_paths:
                        seen_paths.add(result_key)
                        aggregated.append(result)
            else:
                aggregated.append(
                    SearchResult(
                        requested_name=requested_name,
                        file_name="",
                        path="",
                        extension="",
                        size_bytes=0,
                        modified_time="",
                        match_type=mode,
                        status="Not found",
                    )
                )

        return aggregated

    @staticmethod
    def search_roe(source_dir: str, requested_names: list[str]) -> list[SearchResult]:
        all_files = iter_files(source_dir)
        results: list[SearchResult] = []

        for requested_name in requested_names:
            roe_name = normalize_roe_name(requested_name)
            base_name = os.path.splitext(roe_name)[0]
            matches = find_roe_variants(base_name, all_files)

            if not matches:
                results.append(
                    SearchResult(
                        requested_name=requested_name,
                        file_name="",
                        path="",
                        extension="",
                        size_bytes=0,
                        modified_time="",
                        match_type="ROE",
                        status="Not found",
                    )
                )
                continue

            for index, path in enumerate(matches, start=1):
                file_name = os.path.basename(path)
                size_bytes, modified_time = safe_stat(path)
                gps_status, latitude, longitude = extract_gps_metadata(path)
                results.append(
                    SearchResult(
                        requested_name=requested_name,
                        file_name=file_name,
                        path=path,
                        extension=Path(file_name).suffix.lower(),
                        size_bytes=size_bytes,
                        modified_time=modified_time,
                        match_type=f"ROE variant {index}",
                        status="Found",
                        gps_status=gps_status,
                        latitude=latitude,
                        longitude=longitude,
                    )
                )

        return results


class ResultTree:
    def __init__(self, master: tk.Widget):
        frame = tk.Frame(master)
        frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(frame, columns=SEARCH_COLUMNS, show="headings", selectmode="extended")
        for column in SEARCH_COLUMNS:
            self.tree.heading(column, text=column)
            width = 140
            if column == "Path":
                width = 420
            elif column in {"Request", "File"}:
                width = 180
            self.tree.column(column, width=width, stretch=True)

        y_scroll = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        x_scroll = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        self.results: list[SearchResult] = []

    def set_results(self, results: list[SearchResult]) -> None:
        self.clear()
        self.results = results
        for index, result in enumerate(results):
            values = (
                result.requested_name,
                result.file_name,
                result.path,
                result.extension,
                format_size(result.size_bytes) if result.size_bytes else "",
                result.modified_time,
                result.match_type,
                result.status,
            )
            self.tree.insert("", "end", iid=str(index), values=values)

    def clear(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.results = []

    def selected_results(self) -> list[SearchResult]:
        selected: list[SearchResult] = []
        for item_id in self.tree.selection():
            try:
                selected.append(self.results[int(item_id)])
            except Exception:
                continue
        return selected

    def all_results(self) -> list[SearchResult]:
        return list(self.results)


class FinderApp:
    def __init__(self, master: tk.Tk):
        self.master = master
        self.master.title(APP_TITLE)
        self.master.geometry(WINDOW_SIZE)
        self._set_icon()

        self.search_source_dir = tk.StringVar()
        self.search_query = tk.StringVar()
        self.search_mode = tk.StringVar(value="Exact")

        self.batch_source_dir = tk.StringVar()
        self.batch_file_path = tk.StringVar()
        self.batch_mode = tk.StringVar(value="Exact")
        self.batch_column = tk.StringVar()
        self.batch_df: pd.DataFrame | None = None

        self.copy_destination_dir = tk.StringVar()

        self.roe_source_dir = tk.StringVar()
        self.roe_destination_dir = tk.StringVar()
        self.roe_file_path = tk.StringVar()
        self.roe_names_text = tk.StringVar()
        self.roe_df: pd.DataFrame | None = None
        self.roe_column = tk.StringVar()

        self.gps_input_dir = tk.StringVar()
        self.gps_output_dir = tk.StringVar()

        self.last_results: list[SearchResult] = []
        self._build_ui()

    def _set_icon(self) -> None:
        try:
            icon_image = Image.new("RGBA", (32, 32), (255, 255, 255, 0))
            icon_photo = ImageTk.PhotoImage(icon_image)
            self.master.iconphoto(True, icon_photo)
            self.master._icon_photo = icon_photo
        except Exception:
            pass

    def _build_ui(self) -> None:
        self.notebook = ttk.Notebook(self.master)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.search_tab = tk.Frame(self.notebook)
        self.batch_tab = tk.Frame(self.notebook)
        self.copy_tab = tk.Frame(self.notebook)
        self.roe_tab = tk.Frame(self.notebook)
        self.log_tab = tk.Frame(self.notebook)

        self.notebook.add(self.search_tab, text="Search")
        self.notebook.add(self.batch_tab, text="Batch Excel / CSV")
        self.notebook.add(self.copy_tab, text="Safe Copy")
        self.notebook.add(self.roe_tab, text="Roe Photos / Metadata")
        self.notebook.add(self.log_tab, text="Logs")

        self._build_search_tab()
        self._build_batch_tab()
        self._build_copy_tab()
        self._build_roe_tab()
        self._build_log_tab()

    def _build_search_tab(self) -> None:
        controls = tk.LabelFrame(self.search_tab, text="File search", padx=10, pady=10)
        controls.pack(fill="x", padx=10, pady=10)

        tk.Label(controls, text="Source folder:").grid(row=0, column=0, sticky="w")
        tk.Entry(controls, textvariable=self.search_source_dir, width=85).grid(row=0, column=1, padx=5, pady=4, sticky="we")
        tk.Button(controls, text="Browse...", command=lambda: self.select_directory(self.search_source_dir)).grid(row=0, column=2, padx=5)

        tk.Label(controls, text="Search:").grid(row=1, column=0, sticky="w")
        tk.Entry(controls, textvariable=self.search_query, width=50).grid(row=1, column=1, padx=5, pady=4, sticky="w")
        ttk.Combobox(controls, textvariable=self.search_mode, values=["Exact", "Contains", "Similar"], state="readonly", width=15).grid(row=1, column=2, padx=5, sticky="w")

        button_row = tk.Frame(controls)
        button_row.grid(row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))
        tk.Button(button_row, text="Search", bg="#1f6aa5", fg="white", command=self.run_search).pack(side="left", padx=(0, 6))
        tk.Button(button_row, text="Open selected file", command=lambda: self.open_selected_from_tree(self.search_results)).pack(side="left", padx=6)
        tk.Button(button_row, text="Open folder", command=lambda: self.open_folder_for_selected(self.search_results)).pack(side="left", padx=6)
        tk.Button(button_row, text="Send to Safe Copy", command=lambda: self.push_results_to_copy(self.search_results.all_results())).pack(side="left", padx=6)

        results_frame = tk.LabelFrame(self.search_tab, text="Search results", padx=10, pady=10)
        results_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.search_results = ResultTree(results_frame)

    def _build_batch_tab(self) -> None:
        controls = tk.LabelFrame(self.batch_tab, text="Batch search from Excel / CSV", padx=10, pady=10)
        controls.pack(fill="x", padx=10, pady=10)

        tk.Label(controls, text="Source folder:").grid(row=0, column=0, sticky="w")
        tk.Entry(controls, textvariable=self.batch_source_dir, width=85).grid(row=0, column=1, padx=5, pady=4, sticky="we")
        tk.Button(controls, text="Browse...", command=lambda: self.select_directory(self.batch_source_dir)).grid(row=0, column=2, padx=5)

        tk.Label(controls, text="List file:").grid(row=1, column=0, sticky="w")
        tk.Entry(controls, textvariable=self.batch_file_path, width=85).grid(row=1, column=1, padx=5, pady=4, sticky="we")
        tk.Button(controls, text="Load...", command=self.load_batch_file).grid(row=1, column=2, padx=5)

        tk.Label(controls, text="Name column:").grid(row=2, column=0, sticky="w")
        self.batch_column_combo = ttk.Combobox(controls, textvariable=self.batch_column, state="readonly", width=35)
        self.batch_column_combo.grid(row=2, column=1, padx=5, pady=4, sticky="w")
        ttk.Combobox(controls, textvariable=self.batch_mode, values=["Exact", "Contains", "Similar"], state="readonly", width=15).grid(row=2, column=2, padx=5, sticky="w")

        button_row = tk.Frame(controls)
        button_row.grid(row=3, column=0, columnspan=3, sticky="w", pady=(8, 0))
        tk.Button(button_row, text="Search list", bg="#1f6aa5", fg="white", command=self.run_batch_search).pack(side="left", padx=(0, 6))
        tk.Button(button_row, text="Export results to CSV", command=lambda: self.export_results_csv(self.batch_results.all_results())).pack(side="left", padx=6)
        tk.Button(button_row, text="Send to Safe Copy", command=lambda: self.push_results_to_copy(self.batch_results.all_results())).pack(side="left", padx=6)

        results_frame = tk.LabelFrame(self.batch_tab, text="Batch results", padx=10, pady=10)
        results_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.batch_results = ResultTree(results_frame)

    def _build_copy_tab(self) -> None:
        controls = tk.LabelFrame(self.copy_tab, text="Safe selective copy", padx=10, pady=10)
        controls.pack(fill="x", padx=10, pady=10)

        tk.Label(controls, text="Destination folder:").grid(row=0, column=0, sticky="w")
        tk.Entry(controls, textvariable=self.copy_destination_dir, width=85).grid(row=0, column=1, padx=5, pady=4, sticky="we")
        tk.Button(controls, text="Browse...", command=lambda: self.select_directory(self.copy_destination_dir)).grid(row=0, column=2, padx=5)

        button_row = tk.Frame(controls)
        button_row.grid(row=1, column=0, columnspan=3, sticky="w", pady=(8, 0))
        tk.Button(button_row, text="Copy selected", bg="#2d7d46", fg="white", command=self.copy_selected_results).pack(side="left", padx=(0, 6))
        tk.Button(button_row, text="Copy all found results", command=self.copy_all_results).pack(side="left", padx=6)
        tk.Button(button_row, text="Open selected file", command=lambda: self.open_selected_from_tree(self.copy_results)).pack(side="left", padx=6)
        tk.Button(button_row, text="Open folder", command=lambda: self.open_folder_for_selected(self.copy_results)).pack(side="left", padx=6)
        tk.Button(button_row, text="Clear list", command=lambda: self.set_copy_results([])).pack(side="left", padx=6)

        info = tk.Label(
            controls,
            text="Safety rule: source files are never modified and duplicate names in the destination are renamed automatically.",
            justify="left",
            fg="#444444",
        )
        info.grid(row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))

        results_frame = tk.LabelFrame(self.copy_tab, text="Items ready to copy", padx=10, pady=10)
        results_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.copy_results = ResultTree(results_frame)

    def _build_roe_tab(self) -> None:
        top_frame = tk.Frame(self.roe_tab)
        top_frame.pack(fill="both", expand=True, padx=10, pady=10)

        roe_frame = tk.LabelFrame(top_frame, text="Roe search and copy", padx=10, pady=10)
        roe_frame.pack(fill="x", pady=(0, 10))

        tk.Label(roe_frame, text="Source folder:").grid(row=0, column=0, sticky="w")
        tk.Entry(roe_frame, textvariable=self.roe_source_dir, width=85).grid(row=0, column=1, padx=5, pady=4, sticky="we")
        tk.Button(roe_frame, text="Browse...", command=lambda: self.select_directory(self.roe_source_dir)).grid(row=0, column=2, padx=5)

        tk.Label(roe_frame, text="Safe copy folder:").grid(row=1, column=0, sticky="w")
        tk.Entry(roe_frame, textvariable=self.roe_destination_dir, width=85).grid(row=1, column=1, padx=5, pady=4, sticky="we")
        tk.Button(roe_frame, text="Browse...", command=lambda: self.select_directory(self.roe_destination_dir)).grid(row=1, column=2, padx=5)

        tk.Label(roe_frame, text="Roe list file:").grid(row=2, column=0, sticky="w")
        tk.Entry(roe_frame, textvariable=self.roe_file_path, width=85).grid(row=2, column=1, padx=5, pady=4, sticky="we")
        tk.Button(roe_frame, text="Load...", command=self.load_roe_file).grid(row=2, column=2, padx=5)

        tk.Label(roe_frame, text="Name column:").grid(row=3, column=0, sticky="w")
        self.roe_column_combo = ttk.Combobox(roe_frame, textvariable=self.roe_column, state="readonly", width=35)
        self.roe_column_combo.grid(row=3, column=1, padx=5, pady=4, sticky="w")

        tk.Label(roe_frame, text="Manual names (comma-separated):").grid(row=4, column=0, sticky="nw")
        tk.Entry(roe_frame, textvariable=self.roe_names_text, width=85).grid(row=4, column=1, padx=5, pady=4, sticky="we")

        roe_buttons = tk.Frame(roe_frame)
        roe_buttons.grid(row=5, column=0, columnspan=3, sticky="w", pady=(8, 0))
        tk.Button(roe_buttons, text="Search Roe", bg="#1f6aa5", fg="white", command=self.run_roe_search).pack(side="left", padx=(0, 6))
        tk.Button(roe_buttons, text="Copy Roe results", command=self.copy_all_roe_results).pack(side="left", padx=6)
        tk.Button(roe_buttons, text="Export Roe log", command=lambda: self.export_results_excel(self.roe_results.all_results(), include_gps=True)).pack(side="left", padx=6)

        roe_results_frame = tk.LabelFrame(top_frame, text="Roe results", padx=10, pady=10)
        roe_results_frame.pack(fill="both", expand=True, pady=(0, 10))
        self.roe_results = ResultTree(roe_results_frame)

        gps_frame = tk.LabelFrame(top_frame, text="GPS metadata check on copies", padx=10, pady=10)
        gps_frame.pack(fill="x")

        tk.Label(gps_frame, text="Input photo folder:").grid(row=0, column=0, sticky="w")
        tk.Entry(gps_frame, textvariable=self.gps_input_dir, width=85).grid(row=0, column=1, padx=5, pady=4, sticky="we")
        tk.Button(gps_frame, text="Browse...", command=lambda: self.select_directory(self.gps_input_dir)).grid(row=0, column=2, padx=5)

        tk.Label(gps_frame, text="Output folder:").grid(row=1, column=0, sticky="w")
        tk.Entry(gps_frame, textvariable=self.gps_output_dir, width=85).grid(row=1, column=1, padx=5, pady=4, sticky="we")
        tk.Button(gps_frame, text="Browse...", command=lambda: self.select_directory(self.gps_output_dir)).grid(row=1, column=2, padx=5)

        gps_buttons = tk.Frame(gps_frame)
        gps_buttons.grid(row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))
        tk.Button(gps_buttons, text="Run GPS check", bg="#1f6aa5", fg="white", command=self.run_gps_check).pack(side="left", padx=(0, 6))
        tk.Label(
            gps_buttons,
            text="Originals are never modified: any GPS metadata is written only to output copies.",
            fg="#444444",
        ).pack(side="left", padx=10)

    def _build_log_tab(self) -> None:
        self.log_widget = scrolledtext.ScrolledText(self.log_tab, wrap=tk.WORD)
        self.log_widget.pack(fill="both", expand=True, padx=10, pady=10)

    def log(self, message: str) -> None:
        timestamp = pd.Timestamp.now().strftime("%H:%M:%S")
        self.log_widget.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_widget.see(tk.END)
        self.master.update_idletasks()

    def select_directory(self, target_var: tk.StringVar) -> None:
        selected = filedialog.askdirectory()
        if selected:
            target_var.set(selected)

    def select_file(self, target_var: tk.StringVar) -> None:
        selected = filedialog.askopenfilename(
            filetypes=[("Tabular files", "*.xlsx;*.xls;*.csv"), ("All files", "*.*")]
        )
        if selected:
            target_var.set(selected)

    def open_selected_from_tree(self, result_tree: ResultTree) -> None:
        selected = result_tree.selected_results()
        if not selected or not selected[0].path:
            messagebox.showwarning("Open file", "Select a valid file.")
            return
        target_path = selected[0].path
        ext = Path(target_path).suffix.lower()
        if ext in POTENTIALLY_DANGEROUS_EXTENSIONS:
            confirm = messagebox.askyesno(
                "Security Warning",
                f"The selected file '{os.path.basename(target_path)}' is an executable or script ({ext}).\n\n"
                "Opening it will execute code on your system.\n\n"
                "Do you really want to open and run this file?",
                icon="warning",
            )
            if not confirm:
                return
        try:
            os.startfile(target_path)  # type: ignore[attr-defined]
        except Exception as exc:
            messagebox.showerror("Error", f"Unable to open file:\n{exc}")

    def open_folder_for_selected(self, result_tree: ResultTree) -> None:
        selected = result_tree.selected_results()
        if not selected or not selected[0].path:
            messagebox.showwarning("Open folder", "Select a valid file.")
            return
        try:
            os.startfile(os.path.dirname(selected[0].path))  # type: ignore[attr-defined]
        except Exception as exc:
            messagebox.showerror("Error", f"Unable to open folder:\n{exc}")

    def push_results_to_copy(self, results: list[SearchResult]) -> None:
        filtered = [result for result in results if result.path and result.status == "Found"]
        if not filtered:
            messagebox.showwarning("Safe Copy", "There are no valid results to send to the copy tab.")
            return
        self.set_copy_results(filtered)
        self.notebook.select(self.copy_tab)
        self.log(f"Sent {len(filtered)} results to the Safe Copy tab.")

    def set_copy_results(self, results: list[SearchResult]) -> None:
        self.copy_results.set_results(results)

    def validate_source_dir(self, source_dir: str) -> bool:
        if not source_dir or not os.path.isdir(source_dir):
            messagebox.showerror("Error", "Select a valid source folder.")
            return False
        return True

    def run_search(self) -> None:
        source_dir = self.search_source_dir.get().strip()
        query = self.search_query.get().strip()
        mode = self.search_mode.get()

        if not self.validate_source_dir(source_dir):
            return
        if not query:
            messagebox.showerror("Error", "Enter text to search for.")
            return

        self.log(f"Search started in '{source_dir}' using {mode} mode: {query}")
        results = SearchEngine.search_by_query(source_dir, query, mode)
        self.search_results.set_results(results)
        self.last_results = results
        self.log(f"Search completed: {len(results)} results.")

    def load_batch_file(self) -> None:
        self.select_file(self.batch_file_path)
        file_path = self.batch_file_path.get().strip()
        if not file_path:
            return

        try:
            self.batch_df = read_tabular_file(file_path)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return

        columns = [str(column) for column in self.batch_df.columns]
        self.batch_column_combo["values"] = columns
        if columns:
            self.batch_column.set(columns[0])
        self.log(f"Loaded batch file: {file_path}")

    def extract_names_from_dataframe(self, dataframe: pd.DataFrame | None, column_name: str) -> list[str]:
        if dataframe is None:
            raise ValueError("No tabular file loaded.")
        if column_name not in dataframe.columns:
            raise ValueError("Select a valid column.")
        values = dataframe[column_name].dropna().astype(str).tolist()
        cleaned = [value.strip() for value in values if value.strip()]
        if not cleaned:
            raise ValueError("The selected column does not contain usable values.")
        return cleaned

    def run_batch_search(self) -> None:
        source_dir = self.batch_source_dir.get().strip()
        if not self.validate_source_dir(source_dir):
            return

        try:
            requested_names = self.extract_names_from_dataframe(self.batch_df, self.batch_column.get())
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return

        mode = self.batch_mode.get()
        self.log(f"Batch search started with {len(requested_names)} requests using {mode} mode.")
        results = SearchEngine.search_batch(source_dir, requested_names, mode)
        self.batch_results.set_results(results)
        self.last_results = results
        found = len([item for item in results if item.status == "Found"])
        missing = len([item for item in results if item.status != "Found"])
        self.log(f"Batch search completed: {found} found, {missing} missing.")

    def copy_selected_results(self) -> None:
        selected = self.copy_results.selected_results()
        if not selected:
            messagebox.showwarning("Safe Copy", "Select at least one result to copy.")
            return
        self.perform_safe_copy(selected)

    def copy_all_results(self) -> None:
        all_results = [result for result in self.copy_results.all_results() if result.path and result.status == "Found"]
        if not all_results:
            messagebox.showwarning("Safe Copy", "There are no found results to copy.")
            return
        self.perform_safe_copy(all_results)

    def perform_safe_copy(self, results: list[SearchResult]) -> None:
        destination_dir = self.copy_destination_dir.get().strip()
        if not destination_dir:
            messagebox.showerror("Error", "Select a destination folder.")
            return

        copied = 0
        for result in results:
            try:
                copied_path = safe_copy_file(result.path, destination_dir)
                copied += 1
                self.log(f"Safely copied: {result.path} -> {copied_path}")
            except Exception as exc:
                self.log(f"COPY ERROR {result.path}: {exc}")

        messagebox.showinfo("Safe Copy", f"Copy complete.\nFiles copied: {copied}")

    def load_roe_file(self) -> None:
        self.select_file(self.roe_file_path)
        file_path = self.roe_file_path.get().strip()
        if not file_path:
            return

        try:
            self.roe_df = read_tabular_file(file_path)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return

        columns = [str(column) for column in self.roe_df.columns]
        self.roe_column_combo["values"] = columns
        if columns:
            self.roe_column.set(columns[0])
        self.log(f"Loaded Roe file: {file_path}")

    def get_roe_requested_names(self) -> list[str]:
        manual_names = [item.strip() for item in self.roe_names_text.get().split(",") if item.strip()]
        file_names: list[str] = []
        if self.roe_df is not None and self.roe_column.get():
            file_names = self.extract_names_from_dataframe(self.roe_df, self.roe_column.get())
        requested = manual_names + file_names
        unique_requested: list[str] = []
        seen: set[str] = set()
        for item in requested:
            normalized = normalize_name(item)
            if normalized not in seen:
                seen.add(normalized)
                unique_requested.append(item)
        return unique_requested

    def run_roe_search(self) -> None:
        source_dir = self.roe_source_dir.get().strip()
        if not self.validate_source_dir(source_dir):
            return

        try:
            requested_names = self.get_roe_requested_names()
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return

        if not requested_names:
            messagebox.showerror("Error", "Enter Roe names manually or load an XLSX/XLS/CSV file.")
            return

        self.log(f"Roe search started with {len(requested_names)} requests.")
        results = SearchEngine.search_roe(source_dir, requested_names)
        self.roe_results.set_results(results)
        self.log(f"Roe search completed: {len([item for item in results if item.path])} matches.")

    def copy_all_roe_results(self) -> None:
        destination_dir = self.roe_destination_dir.get().strip()
        if not destination_dir:
            messagebox.showerror("Error", "Select a Roe destination folder.")
            return

        results = [result for result in self.roe_results.all_results() if result.path and result.status == "Found"]
        if not results:
            messagebox.showwarning("Roe", "There are no Roe results to copy.")
            return

        copied = 0
        for result in results:
            try:
                copied_path = safe_copy_file(result.path, destination_dir)
                copied += 1
                self.log(f"Copied Roe file: {result.path} -> {copied_path}")
            except Exception as exc:
                self.log(f"ROE COPY ERROR {result.path}: {exc}")

        messagebox.showinfo("Roe", f"Roe copy complete.\nFiles copied: {copied}")

    def export_results_csv(self, results: list[SearchResult]) -> None:
        if not results:
            messagebox.showwarning("Export", "There are no results to export.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
        )
        if not file_path:
            return

        with open(file_path, "w", newline="", encoding="utf-8-sig") as csv_file:
            writer = csv.writer(csv_file)
            writer.writerow(["Request", "File", "Path", "Extension", "Size", "Modified", "Match", "Status"])
            for result in results:
                writer.writerow(
                    [
                        sanitize_spreadsheet_value(result.requested_name),
                        sanitize_spreadsheet_value(result.file_name),
                        sanitize_spreadsheet_value(result.path),
                        sanitize_spreadsheet_value(result.extension),
                        result.size_bytes,
                        sanitize_spreadsheet_value(result.modified_time),
                        sanitize_spreadsheet_value(result.match_type),
                        sanitize_spreadsheet_value(result.status),
                    ]
                )
        self.log(f"Exported CSV results: {file_path}")

    def export_results_excel(self, results: list[SearchResult], include_gps: bool = False) -> None:
        if not results:
            messagebox.showwarning("Export", "There are no results to export.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not file_path:
            return

        rows: list[dict[str, object]] = []
        for result in results:
            row: dict[str, object] = {
                "Requested name": sanitize_spreadsheet_value(result.requested_name),
                "File": sanitize_spreadsheet_value(result.file_name),
                "Path": sanitize_spreadsheet_value(result.path),
                "Extension": sanitize_spreadsheet_value(result.extension),
                "Size bytes": result.size_bytes,
                "Modified time": sanitize_spreadsheet_value(result.modified_time),
                "Match": sanitize_spreadsheet_value(result.match_type),
                "Status": sanitize_spreadsheet_value(result.status),
            }
            if include_gps:
                row["GPS"] = sanitize_spreadsheet_value(result.gps_status)
                row["Latitude"] = sanitize_spreadsheet_value(result.latitude)
                row["Longitude"] = sanitize_spreadsheet_value(result.longitude)
            rows.append(row)

        dataframe = pd.DataFrame(rows)
        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            dataframe.to_excel(writer, index=False, sheet_name="Results")
            worksheet = writer.sheets["Results"]
            for cell in worksheet["C"][1:]:
                cell.alignment = Alignment(wrapText=True)

        self.log(f"Exported Excel results: {file_path}")

    def prompt_gps_coordinates(
        self,
        file_name: str,
        current_index: int,
        total_count: int,
        last_latitude: str,
        last_longitude: str,
    ) -> tuple[str, str, bool] | None:
        result: dict[str, object] = {"value": None}
        popup = tk.Toplevel(self.master)
        popup.title(f"GPS coordinates {current_index}/{total_count}")
        popup.grab_set()

        tk.Label(popup, text=f"File {current_index} of {total_count}\n{file_name}").pack(padx=15, pady=10)
        tk.Label(popup, text="Latitude:").pack(anchor="w", padx=15)
        lat_var = tk.StringVar(value=last_latitude)
        tk.Entry(popup, textvariable=lat_var, width=30).pack(padx=15, pady=4)
        tk.Label(popup, text="Longitude:").pack(anchor="w", padx=15)
        lon_var = tk.StringVar(value=last_longitude)
        tk.Entry(popup, textvariable=lon_var, width=30).pack(padx=15, pady=4)

        button_frame = tk.Frame(popup)
        button_frame.pack(pady=12)

        def submit(apply_all: bool) -> None:
            latitude = lat_var.get().strip()
            longitude = lon_var.get().strip()
            try:
                float(latitude)
                float(longitude)
            except ValueError:
                messagebox.showerror("Error", "Enter valid numeric coordinates.", parent=popup)
                return
            result["value"] = (latitude, longitude, apply_all)
            popup.destroy()

        tk.Button(button_frame, text="Apply and continue", command=lambda: submit(False)).pack(side="left", padx=5)
        tk.Button(button_frame, text="Apply to all", command=lambda: submit(True)).pack(side="left", padx=5)
        tk.Button(button_frame, text="Cancel", command=popup.destroy).pack(side="left", padx=5)

        popup.wait_window()
        return result["value"]  # type: ignore[return-value]

    def run_gps_check(self) -> None:
        input_dir = self.gps_input_dir.get().strip()
        output_dir = self.gps_output_dir.get().strip()

        if not input_dir or not os.path.isdir(input_dir):
            messagebox.showerror("Error", "Select a valid input photo folder.")
            return
        if not output_dir:
            messagebox.showerror("Error", "Select a valid output folder.")
            return

        os.makedirs(output_dir, exist_ok=True)
        all_files = iter_files(input_dir)
        records: list[GpsProcessRecord] = []
        temp_paths: list[str] = []
        jpeg_without_gps: list[GpsProcessRecord] = []
        non_jpeg_without_gps: list[str] = []

        self.log(f"GPS check started on {len(all_files)} files.")

        try:
            for path in all_files:
                gps_status, _, _ = extract_gps_metadata(path)
                extension = Path(path).suffix.lower()
                if gps_status == "SI":
                    copied_path = safe_copy_file(path, output_dir)
                    records.append(
                        GpsProcessRecord(
                            source_name=os.path.basename(path),
                            output_name=os.path.basename(copied_path),
                            source_path=path,
                            working_path=copied_path,
                            inserted_manually=False,
                            converted_to_jpeg=False,
                        )
                    )
                    self.log(f"Copied photo with existing GPS metadata: {path}")
                elif extension in JPEG_EXTENSIONS:
                    copied_path = safe_copy_file(path, output_dir)
                    jpeg_without_gps.append(
                        GpsProcessRecord(
                            source_name=os.path.basename(path),
                            output_name=os.path.basename(copied_path),
                            source_path=path,
                            working_path=copied_path,
                            inserted_manually=False,
                            converted_to_jpeg=False,
                        )
                    )
                else:
                    non_jpeg_without_gps.append(path)

            if non_jpeg_without_gps:
                convert_choice = messagebox.askyesno(
                    "JPEG conversion",
                    (
                        f"Found {len(non_jpeg_without_gps)} non-JPEG files without GPS metadata.\n\n"
                        "Do you want to convert them to JPEG so coordinates can be added to the copies?"
                    ),
                )
                if convert_choice:
                    for path in non_jpeg_without_gps:
                        try:
                            temp_jpeg = convert_to_temp_jpeg(path)
                            temp_paths.append(temp_jpeg)
                            copied_path = safe_copy_file(
                                temp_jpeg,
                                output_dir,
                                output_name=f"{Path(path).stem}.jpg",
                            )
                            jpeg_without_gps.append(
                                GpsProcessRecord(
                                    source_name=os.path.basename(path),
                                    output_name=os.path.basename(copied_path),
                                    source_path=path,
                                    working_path=copied_path,
                                    inserted_manually=False,
                                    converted_to_jpeg=True,
                                )
                            )
                            self.log(f"Converted and copied to JPEG: {path}")
                        except Exception as exc:
                            self.log(f"JPEG CONVERSION ERROR {path}: {exc}")
                else:
                    for path in non_jpeg_without_gps:
                        copied_path = safe_copy_file(path, output_dir)
                        records.append(
                            GpsProcessRecord(
                                source_name=os.path.basename(path),
                                output_name=os.path.basename(copied_path),
                                source_path=path,
                                working_path=copied_path,
                                inserted_manually=False,
                                converted_to_jpeg=False,
                            )
                        )
                        self.log(f"Copied non-JPEG file without GPS metadata, unchanged: {path}")

            last_latitude = ""
            last_longitude = ""
            apply_all_values: tuple[str, str] | None = None

            for index, record in enumerate(jpeg_without_gps, start=1):
                if apply_all_values is None:
                    prompt = self.prompt_gps_coordinates(
                        record.output_name,
                        index,
                        len(jpeg_without_gps),
                        last_latitude,
                        last_longitude,
                    )
                    if prompt is None:
                        self.log("GPS entry interrupted by user.")
                        records.extend(jpeg_without_gps[index - 1 :])
                        break
                    latitude, longitude, apply_all = prompt
                    last_latitude = latitude
                    last_longitude = longitude
                    if apply_all:
                        apply_all_values = (latitude, longitude)
                else:
                    latitude, longitude = apply_all_values

                try:
                    write_gps_metadata(record.working_path, float(latitude), float(longitude))
                    record.inserted_manually = True
                    self.log(f"GPS coordinates written to copy: {record.working_path}")
                except Exception as exc:
                    self.log(f"GPS WRITE ERROR {record.working_path}: {exc}")
                records.append(record)

            log_rows: list[dict[str, str]] = []
            manual_rows: list[dict[str, str]] = []
            for record in records:
                gps_status, latitude, longitude = extract_gps_metadata(record.working_path)
                row = {
                    "Original name": record.source_name,
                    "Output name": record.output_name,
                    "Source path": record.source_path,
                    "Output path": record.working_path,
                    "GPS": gps_status,
                    "Latitude": latitude,
                    "Longitude": longitude,
                    "Inserted manually": "YES" if record.inserted_manually else "NO",
                    "Converted to JPEG": "YES" if record.converted_to_jpeg else "NO",
                }
                log_rows.append(row)
                if record.inserted_manually:
                    manual_rows.append(
                        {
                            "Output name": record.output_name,
                            "Latitude": latitude,
                            "Longitude": longitude,
                        }
                    )

            log_path = os.path.join(
                output_dir,
                f"gps_check_log_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            )
            with pd.ExcelWriter(log_path, engine="openpyxl") as writer:
                pd.DataFrame(log_rows).to_excel(writer, index=False, sheet_name="Full log")
                pd.DataFrame(manual_rows).to_excel(writer, index=False, sheet_name="Manual GPS")

            self.log(f"GPS check completed. Log saved to {log_path}")
            messagebox.showinfo("GPS Check", f"Operation completed.\nLog: {log_path}")
        finally:
            for temp_path in temp_paths:
                try:
                    os.remove(temp_path)
                except OSError:
                    pass


def main() -> None:
    root = tk.Tk()
    app = FinderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
