"""Local-first file manager. Python 3.10+; standard-library core plus local document adapters."""
from __future__ import annotations

import json
import base64
import difflib
import io
import keyword
import logging
from logging.handlers import RotatingFileHandler
import math
import os
import queue
import re
import shutil
import site
import sqlite3
import subprocess
import sys
import threading
import time
import tempfile
import tokenize
import urllib.request
import uuid
import zipfile
from datetime import datetime
from html.parser import HTMLParser
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
import tkinter as tk
import xml.etree.ElementTree as ET

APP_NAME = "Локальный файловый менеджер"
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".bmp"}
SUPPORTED = {".txt", ".md", ".rst", ".csv", ".tsv", ".json", ".xml", ".html", ".htm", ".py", ".js", ".ts", ".css", ".log", ".ini", ".yaml", ".yml", ".toml", ".rtf", ".docx", ".xlsx", ".xlsm", ".xls", ".xlsb", ".pptx", ".odt", ".ods", ".odp", ".pdf", ".png", ".jpg", ".jpeg", ".webp", ".bmp", ".sql", ".sh", ".ps1", ".bat", ".cs", ".java", ".go", ".rs", ".c", ".h", ".cpp", ".hpp", ".jsx", ".tsx", ".vue", ".svelte", ".php", ".rb", ".swift", ".kt", ".properties", ".env"}
MAX_TEXT = 500_000
CUDA_RUNTIME_URL = "https://github.com/ggml-org/llama.cpp/releases/download/b12755/llama-b12755-bin-win-cuda-12.4-x64.zip"
CUDA_RUNTIME_DIR = Path(os.environ.get("APPDATA", Path.home())) / "LocalFileManager" / "runtime" / "llama-cpp-b12755"
_INFERENCE_LOCK = threading.RLock()
_LOADED_MODEL = {"path": None, "runtime": None, "gpu": False}
_APP_LOGGER = None
EDITABLE_TEXT = {".txt", ".md", ".rst", ".csv", ".tsv", ".json", ".xml", ".html", ".htm", ".py", ".js", ".ts", ".css", ".log", ".ini", ".yaml", ".yml", ".toml", ".sql", ".sh", ".ps1", ".bat", ".cs", ".java", ".go", ".rs", ".c", ".h", ".cpp", ".hpp", ".jsx", ".tsx", ".vue", ".svelte", ".php", ".rb", ".swift", ".kt", ".properties", ".env"}
LANGUAGE_NAMES = {".py": "Python", ".php": "PHP", ".js": "JavaScript", ".jsx": "JavaScript JSX", ".ts": "TypeScript", ".tsx": "TypeScript JSX", ".vue": "Vue", ".svelte": "Svelte", ".html": "HTML", ".htm": "HTML", ".css": "CSS", ".scss": "SCSS", ".sql": "SQL", ".java": "Java", ".cs": "C#", ".go": "Go", ".rs": "Rust", ".c": "C", ".h": "C header", ".cpp": "C++", ".hpp": "C++ header", ".rb": "Ruby", ".swift": "Swift", ".kt": "Kotlin", ".sh": "Shell", ".ps1": "PowerShell", ".json": "JSON", ".xml": "XML", ".yaml": "YAML", ".yml": "YAML", ".toml": "TOML"}
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def app_data() -> Path:
    root = Path(os.environ.get("APPDATA", Path.home())) / "LocalFileManager"
    root.mkdir(parents=True, exist_ok=True)
    return root


def app_logger():
    global _APP_LOGGER
    if _APP_LOGGER is None:
        logger = logging.getLogger("local_file_manager")
        logger.setLevel(logging.INFO)
        logger.propagate = False
        if not logger.handlers:
            handler = RotatingFileHandler(app_data() / "app.log", maxBytes=2_000_000, backupCount=3, encoding="utf-8")
            handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
            logger.addHandler(handler)
        _APP_LOGGER = logger
    return _APP_LOGGER


def log_event(message: str, *, exc_info=False):
    try: app_logger().info(message, exc_info=exc_info)
    except OSError: pass


def read_config() -> dict:
    p = app_data() / "config.json"
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {"folders": [], "model": "qwen2.5:7b", "model_dir": str(Path("J:/OllamaModels"))}


def write_config(config: dict) -> None:
    (app_data() / "config.json").write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")


def clean_html(source: str) -> str:
    class TextOnly(HTMLParser):
        def __init__(self):
            super().__init__(); self.parts = []
        def handle_data(self, data): self.parts.append(data)
    parser = TextOnly()
    parser.feed(source)
    return " ".join(parser.parts)


def xml_text(data: bytes) -> str:
    try:
        root = ET.fromstring(data)
        return " ".join(t.strip() for t in root.itertext() if t and t.strip())
    except ET.ParseError:
        return ""


def load_excel_workbook_for_edit(path: Path):
    """Load a spreadsheet into an editable workbook; legacy formats become XLSX data copies."""
    from openpyxl import Workbook, load_workbook
    ext = path.suffix.lower()
    if ext in {".xlsx", ".xlsm"}:
        return load_workbook(path, data_only=False, keep_vba=ext == ".xlsm")
    output = Workbook()
    if ext == ".xls":
        import xlrd
        source = xlrd.open_workbook(str(path), on_demand=True)
        sheets = source.sheets()
        if not sheets: return output
        output.active.title = sheets[0].name[:31] or "Sheet1"
        targets = [output.active]
        for sheet in sheets[1:]: targets.append(output.create_sheet(sheet.name[:31] or "Sheet"))
        for sheet, target in zip(sheets, targets):
            for row in range(sheet.nrows):
                for col in range(sheet.ncols):
                    cell = sheet.cell(row, col)
                    if cell.ctype != xlrd.XL_CELL_EMPTY:
                        target.cell(row + 1, col + 1).value = cell.value
        source.release_resources()
        return output
    if ext == ".xlsb":
        import pyxlsb
        from openpyxl.utils import get_column_letter
        with pyxlsb.open_workbook(str(path)) as source:
            for index, name in enumerate(source.sheets):
                target = output.active if index == 0 else output.create_sheet()
                target.title = (name[:31] or f"Sheet{index + 1}")
                with source.get_sheet(name) as sheet:
                    for row in sheet.rows():
                        for cell in row:
                            if cell.v is not None:
                                target.cell(cell.r + 1, cell.c + 1).value = cell.v
        return output
    raise ValueError(f"Формат {ext} нельзя редактировать как таблицу")


def excel_workbook_text(path: Path, max_chars=MAX_TEXT) -> str:
    """Create a bounded, cell-addressed local context for spreadsheet chat/editing."""
    from openpyxl.utils import get_column_letter
    ext = path.suffix.lower(); chunks = []; used = 0; nonempty = 0
    if ext in {".xlsx", ".xlsm"}:
        from openpyxl import load_workbook
        workbook = load_workbook(path, read_only=True, data_only=False, keep_vba=ext == ".xlsm")
        try:
            for sheet in workbook.worksheets:
                chunks.append(f"\n## Лист: {sheet.title}")
                max_row, max_col = min(sheet.max_row or 0, 10000), min(sheet.max_column or 0, 100)
                for row in sheet.iter_rows(max_row=max_row, max_col=max_col):
                    for cell in row:
                        if cell.value is None: continue
                        line = f"{cell.coordinate}: {cell.value}"
                        if used + len(line) > max_chars or nonempty >= 10000: return "\n".join(chunks) + "\n[контекст таблицы ограничен]"
                        chunks.append(line); used += len(line); nonempty += 1
        finally: workbook.close()
    elif ext == ".xls":
        import xlrd
        source = xlrd.open_workbook(str(path), on_demand=True)
        try:
            for sheet in source.sheets():
                chunks.append(f"\n## Лист: {sheet.name}")
                for row in range(min(sheet.nrows, 10000)):
                    for col in range(min(sheet.ncols, 100)):
                        value = sheet.cell_value(row, col)
                        if value in (None, ""): continue
                        line = f"{get_column_letter(col + 1)}{row + 1}: {value}"
                        if used + len(line) > max_chars or nonempty >= 10000: return "\n".join(chunks) + "\n[контекст таблицы ограничен]"
                        chunks.append(line); used += len(line); nonempty += 1
        finally: source.release_resources()
    elif ext == ".xlsb":
        import pyxlsb
        with pyxlsb.open_workbook(str(path)) as source:
            for name in source.sheets:
                chunks.append(f"\n## Лист: {name}")
                with source.get_sheet(name) as sheet:
                    for row in sheet.rows():
                        for cell in row:
                            if cell.v is None: continue
                            line = f"{get_column_letter(cell.c + 1)}{cell.r + 1}: {cell.v}"
                            if used + len(line) > max_chars or nonempty >= 10000: return "\n".join(chunks) + "\n[контекст таблицы ограничен]"
                            chunks.append(line); used += len(line); nonempty += 1
    return "\n".join(chunks)[:max_chars]


def validate_excel_operations(path: Path, plan: dict):
    from openpyxl.utils.cell import coordinate_to_tuple
    operations = plan.get("operations") if isinstance(plan, dict) else None
    if operations is None and isinstance(plan, dict) and isinstance(plan.get("cells"), dict):
        operations = []
        for address, entry in plan["cells"].items():
            if not isinstance(address, str) or "!" not in address: raise ValueError("Ячейка в плане должна иметь адрес Лист!A1")
            sheet, cell = address.rsplit("!", 1)
            value = entry.get("value") if isinstance(entry, dict) and set(entry) == {"value"} else entry
            operations.append({"sheet": sheet, "cell": cell, "value": value})
    if not isinstance(operations, list) or not operations or len(operations) > 2000:
        raise ValueError("Нужен список из 1–2000 точечных изменений ячеек")
    workbook = load_excel_workbook_for_edit(path)
    validated = []; seen = set()
    try:
        for operation in operations:
            if not isinstance(operation, dict) or set(operation) != {"sheet", "cell", "value"}:
                raise ValueError("Каждое изменение должно содержать только лист, ячейку и новое значение")
            sheet_name, coordinate, value = operation["sheet"], operation["cell"], operation["value"]
            if not isinstance(sheet_name, str) or sheet_name not in workbook.sheetnames:
                raise ValueError(f"Неизвестный лист: {sheet_name!r}")
            if not isinstance(coordinate, str): raise ValueError("Неверный адрес ячейки")
            row, col = coordinate_to_tuple(coordinate.upper())
            if row > 1_048_576 or col > 16_384: raise ValueError("Адрес ячейки выходит за границы Excel")
            key = (sheet_name, coordinate.upper())
            if key in seen: raise ValueError(f"Повторная правка ячейки {sheet_name}!{coordinate}")
            seen.add(key)
            if isinstance(value, (dict, list)) or value is None or not isinstance(value, (str, int, float, bool)):
                raise ValueError("Ячейка должна получить текст, число или логическое значение")
            if isinstance(value, str) and len(value) > 10000: raise ValueError("Значение ячейки превышает 10 000 символов")
            if isinstance(value, float) and (value != value or abs(value) == float("inf")): raise ValueError("Недопустимое число")
            old = workbook[sheet_name][coordinate.upper()].value
            if workbook[sheet_name][coordinate.upper()].__class__.__name__ == "MergedCell":
                raise ValueError(f"Ячейка {sheet_name}!{coordinate} входит в объединённый диапазон и недоступна для изменения")
            validated.append({"sheet": sheet_name, "cell": coordinate.upper(), "value": value, "old": old})
    finally: workbook.close()
    return validated


def write_excel_operations(source: Path, target: Path, operations: list[dict]):
    workbook = load_excel_workbook_for_edit(source)
    try:
        for operation in operations:
            if operation["sheet"] not in workbook.sheetnames: raise ValueError("Лист изменился после проверки")
            cell = workbook[operation["sheet"]][operation["cell"]]
            if cell.value != operation.get("old"):
                raise ValueError(f"Ячейка {operation['sheet']}!{operation['cell']} изменилась после предпросмотра; создайте правку заново")
            cell.value = operation["value"]
        workbook.save(target)
    finally: workbook.close()


def docx_paragraphs(path: Path):
    """Return stable paragraph ids and text from a Word document's text parts."""
    records = []
    with zipfile.ZipFile(path) as archive:
        names = sorted(n for n in archive.namelist() if re.fullmatch(r"word/(document|header\d+|footer\d+|footnotes|endnotes)\.xml", n))
        for name in names:
            root = ET.fromstring(archive.read(name))
            for index, paragraph in enumerate(root.iter(f"{{{W_NS}}}p")):
                text = "".join(node.text or "" for node in paragraph.iter(f"{{{W_NS}}}t"))
                if text.strip(): records.append({"id": len(records) + 1, "part": name, "index": index, "text": text})
    return records


def write_docx_edits(source: Path, target: Path, changes: list[dict]):
    """Rewrite selected paragraph text while retaining the rest of the DOCX package."""
    by_part = {}
    for change in changes: by_part.setdefault(change["part"], {})[change["index"]] = change.get("new_text", change["text"])
    with zipfile.ZipFile(source, "r") as original, zipfile.ZipFile(target, "w") as output:
        for info in original.infolist():
            data = original.read(info.filename)
            replacements = by_part.get(info.filename)
            if replacements:
                root = ET.fromstring(data); paragraphs = list(root.iter(f"{{{W_NS}}}p"))
                for index, new_text in replacements.items():
                    paragraph = paragraphs[index]
                    nodes = list(paragraph.iter(f"{{{W_NS}}}t"))
                    if nodes:
                        nodes[0].text = new_text
                        if new_text[:1].isspace() or new_text[-1:].isspace(): nodes[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                        for node in nodes[1:]: node.text = ""
                    else:
                        run = ET.SubElement(paragraph, f"{{{W_NS}}}r")
                        text_node = ET.SubElement(run, f"{{{W_NS}}}t"); text_node.text = new_text
                for prefix, quote, uri in re.findall(rb'xmlns(?::([A-Za-z_][\w.-]*))?=(["\'])(.*?)\2', original.read(info.filename)):
                    try: ET.register_namespace(prefix.decode("ascii") if prefix else "", uri.decode("utf-8"))
                    except ValueError: pass
                serialized = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                # Keep namespace declarations referenced only by mc:Ignorable values.
                old_open = re.search(rb"<[A-Za-z_][^>]*>", original.read(info.filename))
                new_open = re.search(rb"<[A-Za-z_][^>]*>", serialized)
                if old_open and new_open:
                    declarations = re.findall(rb'xmlns(?::[A-Za-z_][\w.-]*)?=["\'][^"\']+["\']', old_open.group(0))
                    missing = [d for d in declarations if d not in new_open.group(0)]
                    if missing:
                        insert_at = new_open.end() - (2 if new_open.group(0).endswith(b"/>") else 1)
                        serialized = serialized[:insert_at] + b" " + b" ".join(missing) + serialized[insert_at:]
                data = serialized
            output.writestr(info, data)


def read_editable_text(path: Path):
    raw = path.read_bytes()
    if raw.startswith(b"\xef\xbb\xbf"): encoding = "utf-8-sig"
    elif raw.startswith((b"\xff\xfe", b"\xfe\xff")): encoding = "utf-16"
    else:
        try: raw.decode("utf-8"); encoding = "utf-8"
        except UnicodeDecodeError: encoding = "cp1251"
    return raw.decode(encoding), encoding


def is_plain_text_file(path: Path):
    """Format-independent fallback for user supplied text and source files."""
    try:
        with path.open("rb") as stream: sample = stream.read(8192)
        if not sample: return True
        if sample.startswith((b"\xff\xfe", b"\xfe\xff", b"\xef\xbb\xbf")): return True
        if b"\0" in sample: return False
        try: decoded = sample.decode("utf-8")
        except UnicodeDecodeError: decoded = sample.decode("cp1251")
        visible = sum(char.isprintable() or char in "\r\n\t" for char in decoded)
        return visible / max(1, len(decoded)) >= 0.9
    except OSError: return False


def extract_json_object(response: str):
    cleaned = re.sub(r"<think>.*?</think>", "", response, flags=re.DOTALL).strip()
    start = cleaned.find("{")
    if start < 0: raise ValueError("В ответе нет JSON")
    result, _end = json.JSONDecoder().raw_decode(cleaned[start:])
    return result


def extract_file_search_query(prompt: str):
    """Recognize a natural-language request to find a local file and retain identifying terms."""
    lowered = prompt.casefold()
    intents = ("найди", "найти", "поищи", "отыщи", "где лежит", "где находится", "где найти")
    if not any(word in lowered for word in intents):
        # Users sometimes type Russian while the keyboard is still in English layout.
        en = "qwertyuiop[]asdfghjkl;'zxcvbnm,."
        ru = "йцукенгшщзхъфывапролджэячсмитьбю"
        translated = lowered.translate(str.maketrans(en, ru))
        if any(word in translated for word in intents): lowered = translated
    if not any(word in lowered for word in intents): return None
    local_file_context = ("комп" in lowered or "диск" in lowered or "папк" in lowered or "файл" in lowered or
                          "документ" in lowered or "где лежит" in lowered or "где находится" in lowered)
    if not local_file_context: return None
    stop = {"найди", "найти", "поищи", "отыщи", "где", "лежит", "находится", "у", "меня", "на", "в", "во", "по", "про", "для", "из", "с", "от", "там", "файл", "файлы", "документ", "документы", "моем", "моём", "компе", "компьютере", "диске", "где-то", "гдето", "пожалуйста", "мне", "его", "ее", "её", "тот", "этот", "и", "или", "моя", "моей", "мой", "мое", "моё"}
    terms = [word for word in re.findall(r"[\w.-]+", lowered, flags=re.UNICODE) if len(word) > 1 and word not in stop]
    return terms or []


def extract_text(path: Path) -> str:
    ext = path.suffix.lower()
    try:
        if ext in IMAGE_EXTS: return ""
        if ext in {".xlsx", ".xlsm", ".xls", ".xlsb"}: return excel_workbook_text(path)
        if ext in {".docx", ".pptx", ".odt", ".ods", ".odp"}:
            chunks = []
            with zipfile.ZipFile(path) as archive:
                names = archive.namelist()
                if ext == ".docx": names = [n for n in names if re.fullmatch(r"word/(document|header\d+|footer\d+|footnotes|endnotes)\.xml", n)]
                elif ext == ".pptx": names = [n for n in names if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)]
                else: names = [n for n in names if re.fullmatch(r"content\.xml", n)]
                for name in names[:300]:
                    chunks.append(xml_text(archive.read(name)))
            return "\n".join(chunks)[:MAX_TEXT]
        if ext == ".pdf":
            try:
                from pypdf import PdfReader  # optional, installed by requirements.txt
                return "\n".join((page.extract_text() or "") for page in PdfReader(str(path)).pages[:100])[:MAX_TEXT]
            except ImportError:
                return "[Для извлечения текста из PDF установите компонент pypdf из папки приложения.]"
        raw = path.read_text(encoding="utf-8", errors="replace")
        if ext in {".html", ".htm"}: raw = clean_html(raw)
        elif ext == ".rtf": raw = re.sub(r"\\[a-z]+-?\d* ?|[{}]", " ", raw)
        return raw[:MAX_TEXT]
    except ImportError:
        return "[Для работы с Excel установите компоненты командой: python -m pip install -r requirements.txt]"
    except (OSError, ValueError, zipfile.BadZipFile, KeyError, RuntimeError):
        return ""


class Index:
    def __init__(self):
        self.db = sqlite3.connect(app_data() / "files.sqlite3", check_same_thread=False)
        self.db.create_function("CASEFOLD", 1, search_fold)
        self.lock = threading.Lock()
        with self.db:
            self.db.execute("CREATE TABLE IF NOT EXISTS files(path TEXT PRIMARY KEY, name TEXT, ext TEXT, size INTEGER, mtime REAL, body TEXT)")

    def scan(self, folders: list[str], progress):
        found = 0
        for folder in folders:
            base = Path(folder)
            if not base.is_dir(): continue
            for root, dirs, files in os.walk(base):
                dirs[:] = [d for d in dirs if d not in {"$RECYCLE.BIN", "System Volume Information", ".git", "node_modules", "__pycache__"}]
                for filename in files:
                    path = Path(root) / filename
                    try:
                        st = path.stat()
                        if st.st_size > 100_000_000: continue
                        p = str(path.resolve())
                        with self.lock:
                            row = self.db.execute("SELECT size,mtime FROM files WHERE path=?", (p,)).fetchone()
                        if row == (st.st_size, st.st_mtime): continue
                        body = extract_text(path) if path.suffix.lower() in SUPPORTED else ""
                        with self.lock, self.db:
                            self.db.execute("INSERT OR REPLACE INTO files VALUES(?,?,?,?,?,?)", (p, path.name, path.suffix.lower(), st.st_size, st.st_mtime, body))
                        found += 1
                        if found % 30 == 0: progress(f"Проиндексировано: {found} новых или изменённых файлов")
                    except (OSError, PermissionError): continue
        progress(f"Индексация завершена. Обновлено файлов: {found}")

    def search(self, query: str, limit=300):
        with self.lock:
            if not query.strip():
                rows = self.db.execute("SELECT path,name,ext,size,mtime,body FROM files ORDER BY mtime DESC LIMIT ?", (limit,)).fetchall()
            else:
                words = [w.strip() for w in query.split() if len(w.strip()) > 1][:8]
                clauses = " AND ".join(["(CASEFOLD(name) LIKE ? OR CASEFOLD(body) LIKE ? OR CASEFOLD(path) LIKE ?)"] * len(words))
                args = [v for word in words for v in (f"%{search_fold(word)}%", f"%{search_fold(word)}%", f"%{search_fold(word)}%")]
                rows = self.db.execute(f"SELECT path,name,ext,size,mtime,body FROM files WHERE {clauses} ORDER BY mtime DESC LIMIT ?", (*args, limit)).fetchall() if words else []
        return rows

    def find_relevant(self, terms: list[str], limit=10):
        if not terms: return []
        terms = list(dict.fromkeys(search_word_stem(term) for term in terms))[:12]
        clauses = " OR ".join(["(CASEFOLD(name) LIKE ? OR CASEFOLD(body) LIKE ? OR CASEFOLD(path) LIKE ?)"] * len(terms))
        args = [value for term in terms for value in (f"%{term}%", f"%{term}%", f"%{term}%")]
        with self.lock:
            rows = self.db.execute(f"SELECT path,name,ext,size,mtime,body FROM files WHERE {clauses} LIMIT 2000", args).fetchall()
        ranked = []
        for row in rows:
            path, name, ext, size, mtime, body = row
            filename = search_fold(name); full_path = search_fold(path); content = search_fold(body)
            score = sum((5 if term in filename else 0) + (2 if term in full_path else 0) + (1 if term in content else 0) for term in terms)
            if score: ranked.append((score, mtime, row))
        ranked.sort(key=lambda item: (item[0], item[1]), reverse=True)
        return [item[2] for item in ranked[:limit]]

    def counts(self):
        with self.lock:
            return self.db.execute("SELECT COUNT(*), COALESCE(SUM(size),0) FROM files").fetchone()

    def refresh_path(self, path: Path):
        st = path.stat(); p = str(path.resolve())
        body = extract_text(path) if path.suffix.lower() in SUPPORTED else ""
        with self.lock, self.db:
            self.db.execute("INSERT OR REPLACE INTO files VALUES(?,?,?,?,?,?)", (p, path.name, path.suffix.lower(), st.st_size, st.st_mtime, body))

    def relocate(self, old: Path, new: Path):
        st = new.stat(); p = str(new.resolve())
        body = extract_text(new) if new.suffix.lower() in SUPPORTED else ""
        with self.lock, self.db:
            cursor = self.db.execute("UPDATE files SET path=?,name=?,ext=?,size=?,mtime=?,body=? WHERE path=?", (p, new.name, new.suffix.lower(), st.st_size, st.st_mtime, body, str(old.resolve())))
            if cursor.rowcount == 0:
                self.db.execute("INSERT OR REPLACE INTO files VALUES(?,?,?,?,?,?)", (p, new.name, new.suffix.lower(), st.st_size, st.st_mtime, body))


def discover_model_files(models_dir: str | Path) -> dict[str, Path]:
    """Read GGUF model paths from model manifests or a plain folder of GGUF files."""
    root = Path(models_dir).expanduser()
    found = {}
    if root.is_file() and root.suffix.lower() == ".gguf": return {root.stem: root.resolve()}
    manifest_root = root / "manifests"
    if manifest_root.is_dir():
        for manifest_path in manifest_root.rglob("*"):
            if not manifest_path.is_file(): continue
            try: manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            except (OSError, ValueError): continue
            model_layer = next((layer for layer in manifest.get("layers", []) if layer.get("mediaType") == "application/vnd.ollama.image.model"), None)
            if not model_layer: continue
            digest = str(model_layer.get("digest", ""))
            if not digest.startswith("sha256:"): continue
            blob = root / "blobs" / ("sha256-" + digest.split(":", 1)[1])
            if not blob.is_file(): continue
            try: relative = manifest_path.relative_to(manifest_root).parts
            except ValueError: continue
            if len(relative) >= 3:
                label = "/".join(relative[2:-1]) + ":" + relative[-1]
                found[label] = blob.resolve()
    if root.is_dir():
        for model_path in root.rglob("*.gguf"):
            if model_path.is_file(): found.setdefault(model_path.stem, model_path.resolve())
    return dict(sorted(found.items(), key=lambda item: item[0].casefold()))


def resolve_model_path(models_dir: str | Path, model: str) -> Path:
    root = Path(models_dir).expanduser()
    direct = Path(model).expanduser()
    if direct.is_file() and direct.suffix.lower() == ".gguf": return direct.resolve()
    catalog = discover_model_files(root)
    if model in catalog: return catalog[model]
    try:
        saved = json.loads((app_data() / "config.json").read_text(encoding="utf-8"))
        for item in saved.get("model_files", []):
            candidate = Path(item).expanduser()
            if candidate.is_file() and candidate.suffix.lower() == ".gguf" and candidate.stem == model: return candidate.resolve()
    except (OSError, ValueError, TypeError, AttributeError): pass
    raise FileNotFoundError(f"Модель «{model}» не найдена в папке:\n{root}\nВыберите папку с файлами GGUF или хранилище моделей с каталогами manifests/blobs.")


def local_chat(models_dir: str | Path, model: str, messages: list[dict], timeout=180):
    executable = runtime_executable()
    if not executable:
        raise RuntimeError("Встроенный движок устанавливается автоматически. Подождите минуту и повторите запрос.")
    model_path = resolve_model_path(models_dir, model)
    if any(message.get("_image_paths") for message in messages):
        raise ValueError("Для вложенных изображений выберите модель и формат, поддерживаемые встроенным GGUF-движком.")
    compacted_messages, original_chars, sent_chars = compact_chat_messages(messages)
    device_check = subprocess.run([str(executable), "--list-devices"], capture_output=True, text=True, encoding="utf-8", errors="replace",
                                  timeout=20, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0), env=runtime_environment())
    cuda_match = re.search(r"^\s*(CUDA\d+):", device_check.stdout, re.MULTILINE)
    if device_check.returncode != 0 or not cuda_match:
        detail = (device_check.stderr or device_check.stdout or "llama.cpp не обнаружил CUDA-устройство")[-1200:]
        log_event(f"CUDA_DEVICE_ERROR model={model!r} exit={device_check.returncode} detail={detail!r}")
        raise RuntimeError("CUDA-видеокарта не обнаружена встроенным движком. Проверьте драйвер NVIDIA и журнал приложения.\n" + detail)
    cuda_device = cuda_match.group(1)
    log_event(f"CUDA_DEVICE_SELECTED name={cuda_device} devices={device_check.stdout.strip()!r} gpu_layers=all context_tokens=32768 kv_cache=q8_0")
    system = next((str(item.get("content", "")) for item in compacted_messages if item.get("role") == "system"), "Отвечай на языке пользователя.")
    turns = [item for item in compacted_messages if item.get("role") != "system"]
    prompt = "\n\n".join(("Пользователь" if item.get("role") == "user" else "Ассистент") + ":\n" + str(item.get("content", "")) for item in turns)
    if not prompt: raise ValueError("В сообщении нет текста для модели.")
    log_event(f"LLM_REQUEST_START model={model!r} model_path={str(model_path)!r} device={cuda_device} prompt_chars={len(prompt)} system_chars={len(system)} context_chars={original_chars}->{sent_chars}")
    started = time.monotonic()
    with _INFERENCE_LOCK:
        runtime_root = app_data() / "runtime_requests"; runtime_root.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(prefix="request_", dir=runtime_root) as temp_dir:
            temp = Path(temp_dir); prompt_file = temp / "prompt.txt"; system_file = temp / "system.txt"
            prompt_file.write_text(prompt, encoding="utf-8"); system_file.write_text(system, encoding="utf-8")
            command = [str(executable), "-m", str(model_path), "-f", str(prompt_file), "--system-prompt-file", str(system_file),
                       "-n", "2048", "-dev", cuda_device, "-ngl", "999", "-c", "32768",
                       "-ctk", "q8_0", "-ctv", "q8_0", "-fa", "on", "--batch-size", "512",
                       "--simple-io", "--single-turn", "--no-display-prompt"]
            try:
                result = subprocess.run(command, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=timeout,
                                        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0), env=runtime_environment())
            except subprocess.TimeoutExpired as error:
                log_event(f"LLM_REQUEST_TIMEOUT model={model!r} device={cuda_device} seconds={timeout}")
                raise TimeoutError(f"Локальная модель не завершила ответ за {timeout} секунд.") from error
        diagnostics = "\n".join(line.strip() for line in result.stderr.splitlines() if line.strip())[-5000:]
        if diagnostics: log_event(f"LLM_RUNTIME_DIAGNOSTICS model={model!r} device={cuda_device}\n{diagnostics}")
        if result.returncode:
            detail = (result.stderr or result.stdout)[-2000:]
            log_event(f"LLM_REQUEST_ERROR model={model!r} device={cuda_device} exit={result.returncode} detail={detail!r}")
            raise RuntimeError(f"llama.cpp завершился с кодом {result.returncode}:\n{detail}")
        output = result.stdout
        input_marker = "\n> " + prompt + "\n"
        marker_at = output.rfind(input_marker)
        if marker_at >= 0: output = output[marker_at + len(input_marker):]
        if "\n[ Prompt:" in output: output = output.split("\n[ Prompt:", 1)[0]
        output = output.split("\nExiting...", 1)[0].strip()
        if not output:
            log_event(f"LLM_REQUEST_EMPTY model={model!r} device={cuda_device}")
            raise RuntimeError("Локальная модель вернула пустой ответ.")
        _LOADED_MODEL.update({"path": str(model_path), "runtime": cuda_device, "gpu": True})
        log_event(f"LLM_REQUEST_OK model={model!r} device={cuda_device} elapsed_seconds={time.monotonic()-started:.1f}")
        return output


def compact_chat_messages(messages: list[dict], max_chars=42000):
    """Trim old turns to fit the local context window; never silently truncate current content."""
    copied = [dict(message) for message in messages]
    system = next((item for item in copied if item.get("role") == "system"), {"role": "system", "content": ""})
    turns = [item for item in copied if item.get("role") != "system"]
    original_chars = len(str(system.get("content", ""))) + sum(len(str(item.get("content", ""))) for item in turns)
    if not turns: raise ValueError("В запросе нет сообщения пользователя.")
    newest = turns[-1]
    required = len(str(system.get("content", ""))) + len(str(newest.get("content", "")))
    if required > max_chars:
        raise ValueError(f"Текущий запрос вместе с выбранным файлом занимает {required:,} символов, это больше безопасного лимита контекста ({max_chars:,}). Выберите более короткий файл или разделите задачу на части.")
    kept = [newest]
    budget = max_chars - required
    for message in reversed(turns[:-1]):
        cost = len(str(message.get("content", "")))
        if cost + 2 > budget: break
        kept.insert(0, message); budget -= cost + 2
    omitted = len(turns) - len(kept)
    if omitted:
        system = dict(system)
        system["content"] = str(system.get("content", "")) + f"\n\nДля ограничения контекста пропущены {omitted} самых старых сообщений переписки."
    compacted = [system] + kept
    sent_chars = len(str(system.get("content", ""))) + sum(len(str(item.get("content", ""))) for item in kept)
    return compacted, original_chars, sent_chars


def install_inference_runtime() -> tuple[bool, str]:
    """Download the official llama.cpp CUDA runtime. The app invokes its CLI without a server."""
    log_path = app_data() / "engine_install.log"
    creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    if runtime_executable(): return True, str(CUDA_RUNTIME_DIR)
    CUDA_RUNTIME_DIR.mkdir(parents=True, exist_ok=True)
    download_dir = app_data() / "runtime_downloads"; download_dir.mkdir(parents=True, exist_ok=True)
    archive = download_dir / "llama-runtime.zip"
    try:
        dll_dirs = [Path(package_dir) / "nvidia" for package_dir in site.getsitepackages()]
        has_cuda_dlls = any((root / "cuda_runtime" / "bin" / "cudart64_12.dll").is_file() and (root / "cublas" / "bin" / "cublas64_12.dll").is_file() for root in dll_dirs)
        if not has_cuda_dlls:
            command = [sys.executable, "-m", "pip", "install", "--only-binary=:all:", "nvidia-cuda-runtime-cu12==12.4.127", "nvidia-cublas-cu12==12.4.5.8"]
            with log_path.open("a", encoding="utf-8") as log:
                log.write("\\nInstalling CUDA 12.4 runtime libraries\\n")
                result = subprocess.run(command, stdout=log, stderr=subprocess.STDOUT, creationflags=creation_flags, timeout=1800)
            if result.returncode: raise RuntimeError(f"Не удалось поставить CUDA runtime; см. {log_path}")
        if not archive.exists():
            with log_path.open("a", encoding="utf-8") as log:
                log.write("\\nDownloading official llama.cpp b12755 CUDA 12.4 runtime\\n")
            partial = archive.with_suffix(".part")
            with urllib.request.urlopen(CUDA_RUNTIME_URL, timeout=60) as response, partial.open("wb") as destination:
                shutil.copyfileobj(response, destination)
            partial.replace(archive)
        with zipfile.ZipFile(archive) as package:
            base = CUDA_RUNTIME_DIR.resolve()
            for member in package.infolist():
                destination = (CUDA_RUNTIME_DIR / member.filename).resolve()
                if not destination.is_relative_to(base): raise ValueError("Архив runtime содержит небезопасный путь")
            package.extractall(CUDA_RUNTIME_DIR)
        if runtime_executable():
            try: archive.unlink()
            except OSError: pass
            return True, str(CUDA_RUNTIME_DIR)
        return False, f"В архиве не найден llama-cli.exe. Подробности: {log_path}"
    except (OSError, zipfile.BadZipFile, ValueError, urllib.error.URLError) as error:
        return False, f"{error}. Подробности: {log_path}"


def runtime_executable() -> Path | None:
    executable = CUDA_RUNTIME_DIR / "llama-cli.exe"
    return executable if executable.is_file() else None


def runtime_environment() -> dict[str, str]:
    env = os.environ.copy()
    directories = [str(CUDA_RUNTIME_DIR)]
    for package_dir in site.getsitepackages():
        root = Path(package_dir) / "nvidia"
        if root.is_dir(): directories.extend(str(path) for path in root.glob("**/bin"))
    env["PATH"] = os.pathsep.join(directories + [env.get("PATH", "")])
    return env


def human_size(size):
    value = float(size)
    for unit in ("Б", "КБ", "МБ", "ГБ"):
        if value < 1024 or unit == "ГБ": return f"{value:.1f} {unit}"
        value /= 1024


def search_fold(value):
    return str(value or "").casefold().replace("ё", "е")


def search_word_stem(value):
    """Apply a small suffix stemmer so Russian case endings still match file content."""
    word = search_fold(value)
    for suffix in ("иями", "ями", "ами", "ого", "его", "ыми", "ими", "ому", "ему", "иях", "ах", "ях", "ов", "ев", "ей", "ам", "ям", "ом", "ем", "ой", "ый", "ий", "ые", "ие", "ая", "яя", "ое", "ее", "ую", "юю", "а", "я", "ы", "и", "у", "ю", "е", "о"):
        if word.endswith(suffix) and len(word) - len(suffix) >= 4:
            return word[:-len(suffix)]
    return word


def rounded_points(width, height, radius):
    points = []
    for cx, cy, start in ((width-radius, radius, -90), (width-radius, height-radius, 0), (radius, height-radius, 90), (radius, radius, 180)):
        for step in range(7):
            angle = math.radians(start + step * 15)
            points.extend((cx + radius * math.cos(angle), cy + radius * math.sin(angle)))
    return points


def widget_background(master):
    try: return master.cget("background")
    except tk.TclError:
        try:
            name=master.cget("style") or master.winfo_class()
            return ttk.Style(master).lookup(name,"background") or "#f7f7f8"
        except tk.TclError: return "#f7f7f8"


class RoundedEntry(tk.Canvas):
    """Rounded border and inset ttk field, with the familiar Entry API."""
    def __init__(self, master, **entry_options):
        super().__init__(master, height=42, background=widget_background(master), bd=0, highlightthickness=0)
        self.entry = ttk.Entry(self, style="Rounded.TEntry", **entry_options)
        self.entry_item = self.create_window(12, 8, anchor="nw", window=self.entry, width=300, height=26)
        super().bind("<Configure>", self._resize)
        self._draw()

    def _draw(self):
        width=max(80,self.winfo_width()); height=max(36,self.winfo_height())
        super().delete("surface"); self.create_polygon(*rounded_points(width,height,13), smooth=True, splinesteps=18,
            fill="#ffffff", outline="#d1d5db", width=1, tags="surface"); self.tag_lower("surface")
        self.coords(self.entry_item,12,8); self.itemconfigure(self.entry_item,width=max(50,width-24),height=max(20,height-16))

    def _resize(self, _event=None): self._draw()
    def bind(self, sequence=None, func=None, add=None):
        if hasattr(self, "entry"): return self.entry.bind(sequence, func, add)
        return super().bind(sequence, func, add)
    def focus_set(self): self.entry.focus_set()
    def get(self): return self.entry.get()
    def insert(self, index, value): return self.entry.insert(index, value)
    def delete(self, first, last=None): return self.entry.delete(first, last)


class RoundedButton(tk.Canvas):
    """Small, consistent rounded Tk button for a ChatGPT-like desktop surface."""
    def __init__(self, master, text, command=None, style="TButton", state="normal", **kwargs):
        self.label = text; self.command = command; self.variant = style; self.button_state = state
        super().__init__(master, height=42, width=max(88, len(text) * 8 + 34), bd=0, highlightthickness=0,
                         background=self._parent_color(master), **kwargs)
        self.bind("<Configure>", lambda _event: self._draw())
        self.bind("<Enter>", self._hover); self.bind("<Leave>", self._leave)
        self.bind("<ButtonRelease-1>", self._click)
        self.bind("<Return>", self._click); self.bind("<space>", self._click); self.configure(takefocus=True)
        self.bind("<FocusIn>", lambda _e: (setattr(self,"focused",True),self._draw()))
        self.bind("<FocusOut>", lambda _e: (setattr(self,"focused",False),self._draw()))
        self._draw()

    @staticmethod
    def _parent_color(master): return widget_background(master)

    def _colors(self):
        if self.button_state == "disabled": return "#e9ecef", "#9ca3af", "#e9ecef"
        if self.variant == "Accent.TButton":
            return ("#0e8f70", "#ffffff", "#0e8f70") if getattr(self, "hovered", False) else ("#10a37f", "#ffffff", "#10a37f")
        fill = "#eef0f2" if getattr(self, "hovered", False) else "#ffffff"
        return fill, "#202123", "#10a37f" if getattr(self,"focused",False) else "#d9dce1"

    def _draw(self):
        if not self.winfo_exists(): return
        self.delete("all"); width = max(self.winfo_width(), 60); height = max(self.winfo_height(), 36)
        fill, foreground, outline = self._colors(); radius = min(15, height // 2)
        self.create_polygon(*rounded_points(width,height,radius), smooth=True, splinesteps=18, fill=fill, outline=outline, width=1)
        self.create_text(width / 2, height / 2, text=self.label, fill=foreground, font=("Segoe UI Semibold", 9), width=max(40, width - 22))

    def _hover(self, _event=None): self.hovered = True; self._draw()
    def _leave(self, _event=None): self.hovered = False; self._draw()
    def _click(self, _event=None):
        if self.button_state != "disabled" and self.command: self.command()

    def configure(self, cnf=None, **kwargs):
        if isinstance(cnf, dict): kwargs = {**cnf, **kwargs}
        if "text" in kwargs: self.label = kwargs.pop("text")
        if "command" in kwargs: self.command = kwargs.pop("command")
        if "style" in kwargs: self.variant = kwargs.pop("style")
        if "state" in kwargs: self.button_state = kwargs.pop("state")
        if kwargs: super().configure(**kwargs)
        self._draw()

    config = configure


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        log_event("APPLICATION_START")
        self._configure_styles()
        self.title(APP_NAME); self.geometry("1180x760"); self.minsize(900, 600)
        self.config_data = read_config(); self.index = Index(); self.events = queue.Queue(); self.rows = []
        default_models = Path("J:/OllamaModels") if Path("J:/OllamaModels").is_dir() else Path.home() / ".ollama" / "models"
        self.model_dir_var = tk.StringVar(value=self.config_data.get("model_dir") or str(default_models))
        self.model = tk.StringVar(value=self.config_data.get("model", "qwen2.5:7b"))
        self.model_catalog = discover_model_files(self.model_dir_var.get())
        for item in self.config_data.get("model_files", []):
            candidate = Path(item).expanduser()
            if candidate.is_file() and candidate.suffix.lower() == ".gguf": self.model_catalog.setdefault(candidate.stem, candidate.resolve())
        if self.model_catalog and self.model.get() not in self.model_catalog:
            self.model.set(next(iter(self.model_catalog)))
        self.runtime_ready = runtime_executable() is not None
        self.runtime_installing = False
        self.chats = self._read_chats(); self.active_chat_id = None
        self.chat_history = []; self.chat_attachments = []; self.chat_window = None; self.chat_busy = False
        self.chat_mode_var = tk.StringVar(value="Обычный чат"); self.goal_var = tk.StringVar(value="")
        self.chat_edit_path = None
        self.project_file_path = None
        self.chat_working_dir = self.config_data.get("chat_working_dir", "")
        if self.chat_working_dir and not Path(self.chat_working_dir).is_dir(): self.chat_working_dir = ""
        self.chat_project_dirs = [self.chat_working_dir] if self.chat_working_dir else []
        self.last_assistant_answer = ""
        self._build(); self.protocol("WM_DELETE_WINDOW", self._close_app); self._refresh_folders(); self._refresh_results(); self._poll()
        self.after(500, self._ensure_runtime)

    def _configure_styles(self):
        style = ttk.Style(self)
        try: style.theme_use("clam")
        except tk.TclError: pass
        bg, panel, text, muted, accent = "#f7f7f8", "#ffffff", "#202123", "#6b7280", "#10a37f"
        self.configure(background=bg)
        style.configure(".", font=("Segoe UI", 10), background=bg, foreground=text)
        style.configure("TFrame", background=bg)
        style.configure("Surface.TFrame", background=panel)
        style.configure("Sidebar.TFrame", background="#f0f2f5")
        style.configure("Topbar.TFrame", background="#ffffff")
        style.configure("TLabel", background=bg, foreground=text)
        style.configure("Muted.TLabel", background=bg, foreground=muted)
        style.configure("Title.TLabel", background=bg, foreground="#171717", font=("Segoe UI Semibold", 13))
        style.configure("Sidebar.TLabel", background="#f0f2f5", foreground="#202123")
        style.configure("TopbarTitle.TLabel", background="#ffffff", foreground="#202123", font=("Segoe UI Semibold", 11))
        style.configure("TButton", padding=(13, 9), background="#ffffff", foreground=text, borderwidth=1, bordercolor="#d9dce1", focusthickness=0)
        style.map("TButton", background=[("active", "#eef0f2"), ("pressed", "#e4e7eb"), ("disabled", "#f1f2f4")], foreground=[("disabled", "#9ca3af")])
        style.configure("Topbar.TButton", padding=(10, 8), background="#ffffff", foreground="#374151", borderwidth=1, bordercolor="#e1e4e8")
        style.map("Topbar.TButton", background=[("active", "#f0f2f5"), ("pressed", "#e8eaed")])
        style.configure("Accent.TButton", padding=(15, 10), background=accent, foreground="#ffffff", font=("Segoe UI Semibold", 10), borderwidth=0)
        style.map("Accent.TButton", background=[("active", "#0e8f70"), ("pressed", "#0b795e"), ("disabled", "#a4cfc2")])
        style.configure("TEntry", padding=(10, 9), fieldbackground="#ffffff", foreground=text, insertcolor=text, bordercolor="#d1d5db", lightcolor="#d1d5db", darkcolor="#d1d5db")
        style.map("TEntry", fieldbackground=[("focus", "#ffffff")], bordercolor=[("focus", accent)])
        style.configure("Rounded.TEntry", padding=(8, 6), fieldbackground="#ffffff", foreground=text, insertcolor=text, borderwidth=0, relief="flat")
        style.configure("TCombobox", padding=(9, 7), fieldbackground="#ffffff", foreground=text, background="#f3f4f6", arrowcolor="#6b7280", bordercolor="#d1d5db")
        style.map("TCombobox", fieldbackground=[("readonly", "#ffffff"), ("focus", "#ffffff")], foreground=[("readonly", text)], selectbackground=[("readonly", "#ffffff")], selectforeground=[("readonly", text)])
        style.configure("TNotebook", background=bg, borderwidth=0, tabmargins=(4, 5, 4, 0))
        style.configure("TNotebook.Tab", padding=(20, 12), background="#eef0f2", foreground="#6b7280", borderwidth=0)
        style.map("TNotebook.Tab", background=[("selected", "#ffffff"), ("active", "#e7eaee")], foreground=[("selected", "#202123"), ("active", "#202123")])
        style.configure("Treeview", rowheight=31, background="#ffffff", fieldbackground="#ffffff", foreground="#202123", borderwidth=0, font=("Segoe UI", 10))
        style.map("Treeview", background=[("selected", "#dff3ed")], foreground=[("selected", "#145c49")])
        style.configure("Treeview.Heading", padding=(10, 9), background="#f0f2f5", foreground="#4b5563", font=("Segoe UI Semibold", 9), borderwidth=0)
        style.configure("Status.TLabel", background="#ffffff", foreground="#6b7280", padding=(12, 7), font=("Segoe UI", 9))
        style.configure("TSeparator", background="#e5e7eb")

    def _build(self):
        self.columnconfigure(0, weight=1); self.rowconfigure(1, weight=1)
        bar = ttk.Frame(self, style="Topbar.TFrame", padding=(14, 11)); bar.grid(row=0, column=0, sticky="ew"); bar.columnconfigure(2, weight=1)
        ttk.Label(bar, text="Локальные файлы", style="TopbarTitle.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 18))
        ttk.Label(bar, text="Папки:", style="Sidebar.TLabel").grid(row=0, column=1, sticky="w")
        self.folder_var = tk.StringVar(); RoundedEntry(bar, textvariable=self.folder_var, state="readonly").grid(row=0, column=2, sticky="ew", padx=8)
        RoundedButton(bar, text="Добавить папку", command=self.add_folder, style="Topbar.TButton").grid(row=0, column=3, padx=3)
        RoundedButton(bar, text="Убрать", command=self.remove_folder, style="Topbar.TButton").grid(row=0, column=4, padx=3)
        RoundedButton(bar, text="Сканировать", command=self.scan, style="Accent.TButton").grid(row=0, column=5, padx=(4, 0))
        RoundedButton(bar, text="Журнал", command=self.open_app_log, style="Topbar.TButton").grid(row=0, column=6, padx=(5, 0))
        self.main_notebook = ttk.Notebook(self); self.main_notebook.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0,10))
        self.files_tab = ttk.Frame(self.main_notebook); self.chat_tab = ttk.Frame(self.main_notebook); self.cursor_tab = ttk.Frame(self.main_notebook)
        self.main_notebook.add(self.files_tab, text="Файлы и поиск")
        self.main_notebook.add(self.chat_tab, text="Чат")
        self.main_notebook.add(self.cursor_tab, text="Cursor")
        body = ttk.Panedwindow(self.files_tab, orient="horizontal"); body.grid(row=0, column=0, sticky="nsew")
        self.files_tab.columnconfigure(0, weight=1); self.files_tab.rowconfigure(0, weight=1)
        left = ttk.Frame(body, padding=6); left.columnconfigure(0, weight=1); left.rowconfigure(2, weight=1); body.add(left, weight=3)
        search = ttk.Frame(left); search.grid(row=0, column=0, sticky="ew"); search.columnconfigure(0, weight=1)
        self.query = tk.StringVar(); ent = RoundedEntry(search, textvariable=self.query); ent.grid(row=0, column=0, sticky="ew"); ent.bind("<Return>", lambda e: self._refresh_results())
        RoundedButton(search, text="Найти", command=self._refresh_results).grid(row=0, column=1, padx=4)
        ttk.Label(left, text="Ищет в именах файлов и тексте TXT, DOCX, XLSX, PPTX, ODT, CSV, PDF* и других текстовых файлов.", wraplength=500).grid(row=1, column=0, sticky="w", pady=6)
        self.tree = ttk.Treeview(left, columns=("type", "size", "date"), show="tree headings", selectmode="browse")
        self.tree.heading("#0", text="Файл"); self.tree.heading("type", text="Тип"); self.tree.heading("size", text="Размер"); self.tree.heading("date", text="Изменён")
        self.tree.column("#0", width=340); self.tree.column("type", width=75); self.tree.column("size", width=90); self.tree.column("date", width=125)
        self.tree.grid(row=2, column=0, sticky="nsew"); self.tree.bind("<<TreeviewSelect>>", self._show_selected); self.tree.bind("<Double-1>", lambda e: self.open_selected())
        actions = ttk.Frame(left); actions.grid(row=3, column=0, sticky="ew", pady=6)
        RoundedButton(actions, text="Открыть файл", command=self.open_selected).pack(side="left", padx=(0,5))
        RoundedButton(actions, text="Переименовать", command=self.rename_selected).pack(side="left", padx=5)
        RoundedButton(actions, text="Переместить…", command=self.move_selected).pack(side="left", padx=5)
        right = ttk.Frame(body, padding=8); right.columnconfigure(0, weight=1); right.rowconfigure(1, weight=1); body.add(right, weight=2)
        ttk.Label(right, text="Содержимое документа").grid(row=0, column=0, sticky="w")
        self.preview = tk.Text(right, wrap="word", font=("Segoe UI", 10), state="disabled", background="#ffffff", foreground="#374151", insertbackground="#202123", borderwidth=0, padx=12, pady=10); self.preview.grid(row=1, column=0, sticky="nsew", pady=6)
        ttk.Label(right, text="Встроенная локальная модель").grid(row=2, column=0, sticky="w", pady=(6,2))
        model_folder = ttk.Frame(right); model_folder.grid(row=3, column=0, sticky="ew"); model_folder.columnconfigure(0, weight=1)
        RoundedEntry(model_folder, textvariable=self.model_dir_var, state="readonly").grid(row=0, column=0, sticky="ew")
        RoundedButton(model_folder, text="Папка…", command=self.choose_model_dir).grid(row=0, column=1, padx=(4,0))
        model_bar = ttk.Frame(right); model_bar.grid(row=4, column=0, sticky="ew", pady=4); model_bar.columnconfigure(0, weight=1)
        self.model_box = ttk.Combobox(model_bar, textvariable=self.model, values=tuple(self.model_catalog), state="readonly"); self.model_box.grid(row=0, column=0, sticky="ew")
        self.model_box.bind("<<ComboboxSelected>>", self._model_changed)
        RoundedButton(model_bar, text="Добавить GGUF…", command=self._select_model_from_files).grid(row=0, column=1, padx=3)
        RoundedButton(model_bar, text="Обновить", command=self.check_models).grid(row=0, column=2, padx=3)
        ttk.Label(right, text="Модель запускается внутри приложения; отдельный сервер не нужен.", wraplength=370).grid(row=5, column=0, sticky="w", pady=5)
        files_chat = ttk.Frame(body, padding=8); body.add(files_chat, weight=3)
        self._build_files_chat_ui(files_chat)
        self.status = tk.StringVar(value="Добавьте папку и нажмите «Сканировать».")
        status_bar = ttk.Frame(self); status_bar.grid(row=2, column=0, sticky="ew"); status_bar.columnconfigure(0, weight=1)
        ttk.Label(status_bar, textvariable=self.status, style="Status.TLabel", anchor="w").grid(row=0, column=0, sticky="ew")
        self.llm_progress = ttk.Progressbar(status_bar, mode="indeterminate", length=150)
        self.llm_progress.grid(row=0, column=1, sticky="e", padx=12); self.llm_progress.grid_remove()
        self._build_chat_ui(self.chat_tab)
        self._build_cursor_tab()
        self.main_notebook.bind("<<NotebookTabChanged>>", self._main_tab_changed)
        self.main_notebook.select(self.chat_tab)

    def _build_files_chat_ui(self, panel):
        panel.columnconfigure(0, weight=1); panel.rowconfigure(4, weight=1)
        top = ttk.Frame(panel); top.grid(row=0, column=0, sticky="ew"); top.columnconfigure(1, weight=1)
        ttk.Label(top, text="Локальный чат", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        self.files_chat_model_box = ttk.Combobox(top, textvariable=self.model, values=tuple(self.model_catalog), state="readonly", width=18)
        self.files_chat_model_box.grid(row=0, column=1, sticky="e", padx=4); self.files_chat_model_box.bind("<<ComboboxSelected>>", self._model_changed)
        RoundedButton(top, text="Обновить модели", command=self.check_models).grid(row=0, column=2, padx=2)
        chat_row = ttk.Frame(panel); chat_row.grid(row=1, column=0, sticky="ew", pady=(6, 2)); chat_row.columnconfigure(0, weight=1)
        self.files_chat_mode_box = ttk.Combobox(chat_row, textvariable=self.chat_mode_var, values=["Обычный чат", "Работа по цели", "Режим Cursor"], state="readonly", width=18)
        self.files_chat_mode_box.grid(row=0, column=0, sticky="w"); self.files_chat_mode_box.bind("<<ComboboxSelected>>", self._mode_changed)
        self.files_chat_ids = []; self.files_chat_var = tk.StringVar()
        self.files_chat_box = ttk.Combobox(chat_row, textvariable=self.files_chat_var, state="readonly")
        self.files_chat_box.grid(row=0, column=1, sticky="ew", padx=5); self.files_chat_box.bind("<<ComboboxSelected>>", self._select_files_chat)
        RoundedButton(chat_row, text="＋ Чат", command=self.new_chat).grid(row=0, column=2)
        folder_row = ttk.Frame(panel); folder_row.grid(row=2, column=0, sticky="ew", pady=(2, 4)); folder_row.columnconfigure(1, weight=1)
        RoundedButton(folder_row, text="Рабочая папка…", command=self.choose_chat_working_dir).grid(row=0, column=0, padx=(0, 5))
        self.files_chat_folder_label = tk.StringVar(value=self.chat_working_dir or "Папка не выбрана")
        ttk.Label(folder_row, textvariable=self.files_chat_folder_label, style="Muted.TLabel").grid(row=0, column=1, sticky="ew")
        goal_row = ttk.Frame(panel); goal_row.grid(row=3, column=0, sticky="ew", pady=(2, 3)); goal_row.columnconfigure(1, weight=1)
        ttk.Label(goal_row, text="Цель:").grid(row=0, column=0, sticky="w")
        self.files_goal_entry = RoundedEntry(goal_row, textvariable=self.goal_var)
        self.files_goal_entry.grid(row=0, column=1, sticky="ew", padx=5)
        self.files_goal_entry.bind("<FocusOut>", lambda _e: self._save_current_chat()); self.files_goal_entry.bind("<Return>", lambda _e: self._save_current_chat())
        self.files_goal_bar = goal_row
        self.files_chat_display = tk.Text(panel, wrap="word", state="disabled", font=("Segoe UI", 10), padx=12, pady=12, background="#ffffff", foreground="#202123", borderwidth=0, highlightthickness=1, highlightbackground="#e5e7eb")
        self.files_chat_display.grid(row=4, column=0, sticky="nsew", pady=(4, 5))
        scroll = ttk.Scrollbar(panel, orient="vertical", command=self.files_chat_display.yview); scroll.grid(row=4, column=1, sticky="ns", pady=(4, 5)); self.files_chat_display.configure(yscrollcommand=scroll.set)
        self.files_chat_display.tag_configure("user", foreground="#08765b", spacing1=10, spacing3=6, font=("Segoe UI Semibold", 10, "bold")); self.files_chat_display.tag_configure("assistant", foreground="#202123", spacing1=8, spacing3=14); self.files_chat_display.tag_configure("system", foreground="#6b7280", spacing1=8, spacing3=8)
        self.files_attached_label = tk.StringVar(value="Файл не прикреплён")
        ttk.Label(panel, textvariable=self.files_attached_label).grid(row=5, column=0, sticky="w")
        attach_row = ttk.Frame(panel); attach_row.grid(row=6, column=0, sticky="ew", pady=(2, 0))
        RoundedButton(attach_row, text="Выбранный файл", command=self.attach_selected_result).pack(side="left", padx=(0, 4))
        RoundedButton(attach_row, text="＋ Другие файлы…", command=self.attach_chat_file).pack(side="left", padx=4)
        RoundedButton(attach_row, text="Очистить", command=self.clear_chat_files).pack(side="left", padx=4)
        self.files_chat_input = tk.Text(panel, height=4, wrap="word", font=("Segoe UI", 10), padx=10, pady=8, background="#ffffff", foreground="#202123", insertbackground="#202123", borderwidth=0, highlightthickness=1, highlightbackground="#d1d5db", highlightcolor="#10a37f")
        self.files_chat_input.grid(row=7, column=0, sticky="ew", pady=6); self.files_chat_input.bind("<Return>", self._chat_enter)
        controls = ttk.Frame(panel); controls.grid(row=8, column=0, sticky="ew")
        self.files_chat_send_button = RoundedButton(controls, text="Отправить", command=self.send_chat, style="Accent.TButton"); self.files_chat_send_button.pack(side="left", fill="x", expand=True, padx=(0, 3))
        self.files_chat_edit_button = RoundedButton(controls, text="Правки → diff", command=self.prepare_edit); self.files_chat_edit_button.pack(side="left", padx=3)
        self.files_chat_save_button = RoundedButton(controls, text="Сохранить", command=self.save_chat_answer); self.files_chat_save_button.pack(side="left", padx=3)
        ttk.Label(panel, text="Поиск и анализ выполняются локально; правки показываются до сохранения.", style="Muted.TLabel", wraplength=430).grid(row=9, column=0, sticky="w", pady=(6, 0))
        self._refresh_files_chat_selector()

    def _refresh_folders(self): self.folder_var.set("; ".join(self.config_data["folders"]) or "Папки не выбраны")
    def open_app_log(self):
        path = app_data() / "app.log"
        if not path.exists(): path.write_text("Журнал Local File Manager\n", encoding="utf-8")
        try: os.startfile(path)
        except (AttributeError, OSError) as e: messagebox.showerror(APP_NAME, f"Не удалось открыть журнал:\n{e}\n\n{path}")
    def _model_changed(self, _event=None):
        self.config_data.update({"model": self.model.get(), "model_dir": self.model_dir_var.get()})
        write_config(self.config_data)
        self.status.set(f"Выбрана модель {self.model.get()}. Она загрузится при следующем запросе.")
    def _select_model_from_files(self):
        files = filedialog.askopenfilenames(title="Добавить модели GGUF", filetypes=[("Модели GGUF", "*.gguf"), ("Все файлы", "*.*")])
        if not files: return
        # Keep several user-selected files in the catalog without copying them.
        self.model_catalog.update({Path(item).stem: Path(item).resolve() for item in files if Path(item).suffix.lower() == ".gguf"})
        remembered = set(self.config_data.get("model_files", [])); remembered.update(str(Path(item).resolve()) for item in files if Path(item).suffix.lower() == ".gguf")
        self.config_data["model_files"] = sorted(remembered, key=str.casefold); write_config(self.config_data)
        self._set_model_choices(self.model_catalog)
        if files:
            self.model.set(Path(files[0]).stem); self._model_changed()
    def _build_cursor_tab(self):
        tab = self.cursor_tab
        tab.columnconfigure(0, weight=1); tab.rowconfigure(1, weight=1)
        toolbar = ttk.Frame(tab, padding=10); toolbar.grid(row=0, column=0, sticky="ew"); toolbar.columnconfigure(1, weight=1)
        ttk.Label(toolbar, text="Проекты:").grid(row=0, column=0, padx=(0, 6))
        self.cursor_folder_label = tk.StringVar(value=self.chat_working_dir or "Выберите рабочую папку проекта")
        self.cursor_project_box = ttk.Combobox(toolbar, textvariable=self.cursor_folder_label, state="readonly")
        self.cursor_project_box.grid(row=0, column=1, sticky="ew", padx=4); self.cursor_project_box.bind("<<ComboboxSelected>>", self._select_cursor_project)
        RoundedButton(toolbar, text="＋ Проект", command=self.add_cursor_project).grid(row=0, column=2, padx=4)
        ttk.Label(toolbar, text="Чат:").grid(row=0, column=3, padx=(10, 2))
        self.cursor_chat_var = tk.StringVar()
        self.cursor_chat_box = ttk.Combobox(toolbar, textvariable=self.cursor_chat_var, state="readonly", width=24)
        self.cursor_chat_box.grid(row=0, column=4, padx=4); self.cursor_chat_box.bind("<<ComboboxSelected>>", self._select_cursor_chat)
        RoundedButton(toolbar, text="＋ Чат", command=self.new_chat).grid(row=0, column=5, padx=4)
        RoundedButton(toolbar, text="Обновить", command=self.refresh_cursor_files).grid(row=0, column=6, padx=4)
        split = ttk.Panedwindow(tab, orient="horizontal"); split.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 8))
        left = ttk.Frame(split, padding=6); left.columnconfigure(0, weight=1); left.rowconfigure(1, weight=1)
        search_bar = ttk.Frame(left); search_bar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6)); search_bar.columnconfigure(0, weight=1)
        ttk.Label(search_bar, text="Поиск по содержимому проекта").grid(row=0, column=0, columnspan=2, sticky="w")
        self.cursor_search_var = tk.StringVar()
        self.cursor_search_entry = RoundedEntry(search_bar, textvariable=self.cursor_search_var)
        self.cursor_search_entry.grid(row=1, column=0, sticky="ew")
        self.cursor_search_entry.bind("<Return>", lambda _e: self.search_cursor_projects())
        RoundedButton(search_bar, text="Найти", command=self.search_cursor_projects).grid(row=1, column=1, padx=(5, 0))
        self.cursor_tree = ttk.Treeview(left, columns=("size",), show="tree headings", selectmode="browse")
        self.cursor_tree.heading("#0", text="Файлы проекта"); self.cursor_tree.heading("size", text="Размер")
        self.cursor_tree.column("#0", width=300); self.cursor_tree.column("size", width=90)
        self.cursor_tree.grid(row=1, column=0, sticky="nsew"); self.cursor_tree.bind("<<TreeviewSelect>>", self._preview_cursor_file)
        left_scroll = ttk.Scrollbar(left, orient="vertical", command=self.cursor_tree.yview); left_scroll.grid(row=1, column=1, sticky="ns"); self.cursor_tree.configure(yscrollcommand=left_scroll.set)
        self.cursor_search_title = ttk.Label(left, text="Результаты поиска", style="Muted.TLabel")
        self.cursor_search_results = ttk.Treeview(left, columns=("line", "match"), show="tree headings", height=6, selectmode="browse")
        self.cursor_search_results.heading("#0", text="Файл"); self.cursor_search_results.heading("line", text="Строка"); self.cursor_search_results.heading("match", text="Совпадение")
        self.cursor_search_results.column("#0", width=150); self.cursor_search_results.column("line", width=52, stretch=False); self.cursor_search_results.column("match", width=240)
        self.cursor_search_results.bind("<Double-1>", self._open_cursor_search_hit)
        center = ttk.Frame(split, padding=6); center.columnconfigure(0, weight=1); center.rowconfigure(1, weight=1)
        editor_toolbar = ttk.Frame(center); editor_toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 5)); editor_toolbar.columnconfigure(0, weight=1)
        self.cursor_editor_path_var = tk.StringVar(value="Откройте файл проекта")
        ttk.Label(editor_toolbar, textvariable=self.cursor_editor_path_var, style="Title.TLabel").grid(row=0, column=0, sticky="w")
        RoundedButton(editor_toolbar, text="Новый файл", command=self.new_cursor_file).grid(row=0, column=1, padx=2)
        RoundedButton(editor_toolbar, text="Найти", command=self.find_in_cursor_editor).grid(row=0, column=2, padx=2)
        self.cursor_completion_button = RoundedButton(editor_toolbar, text="LLM-вставка", command=self.suggest_cursor_completion)
        self.cursor_completion_button.grid(row=0, column=3, padx=2)
        RoundedButton(editor_toolbar, text="Заменить", command=self.replace_in_cursor_editor).grid(row=0, column=4, padx=2)
        RoundedButton(editor_toolbar, text="К строке", command=self.goto_cursor_line).grid(row=0, column=5, padx=2)
        RoundedButton(editor_toolbar, text="Сохранить", command=self.save_cursor_file, style="Accent.TButton").grid(row=0, column=6, padx=2)
        RoundedButton(editor_toolbar, text="Закрыть вкладку", command=self.close_cursor_editor).grid(row=0, column=7, padx=2)
        self.cursor_editor_tabs = ttk.Notebook(center); self.cursor_editor_tabs.grid(row=1, column=0, sticky="nsew")
        self.cursor_editor_tabs.bind("<<NotebookTabChanged>>", self._cursor_editor_tab_changed)
        self.cursor_editor_tabs.bind("<Control-w>", lambda _e: self.close_cursor_editor())
        editor_actions = ttk.Frame(center); editor_actions.grid(row=2, column=0, sticky="ew", pady=(5, 0))
        self.cursor_editor_status = tk.StringVar(value="Готово · Ctrl+S сохранить · Ctrl+F найти")
        ttk.Label(editor_actions, textvariable=self.cursor_editor_status, style="Muted.TLabel").pack(side="left")
        RoundedButton(editor_actions, text="Проверить PHP", command=self.check_php_syntax).pack(side="right", padx=2)
        RoundedButton(editor_actions, text="Класс / функция", command=self.go_to_symbol).pack(side="right", padx=2)
        RoundedButton(editor_actions, text="Найти использования", command=self.find_cursor_usages).pack(side="right", padx=2)
        RoundedButton(editor_actions, text="Открыть внешне", command=self.open_project_file).pack(side="right", padx=2)
        RoundedButton(editor_actions, text="Прикрепить к чату", command=self.attach_cursor_file).pack(side="right", padx=2)
        self.cursor_open_files = {}; self.cursor_active_path = None; self.cursor_search_hits = {}
        self.cursor_symbol_cache = {}
        right = ttk.Frame(split, padding=8); right.columnconfigure(0, weight=1); right.rowconfigure(1, weight=1)
        cursor_head = ttk.Frame(right); cursor_head.grid(row=0, column=0, sticky="ew"); cursor_head.columnconfigure(1, weight=1)
        ttk.Label(cursor_head, text="Модель:").grid(row=0, column=0)
        self.cursor_model_box = ttk.Combobox(cursor_head, textvariable=self.model, values=tuple(self.model_catalog), state="readonly", width=20)
        self.cursor_model_box.grid(row=0, column=1, sticky="ew", padx=5); self.cursor_model_box.bind("<<ComboboxSelected>>", self._model_changed)
        RoundedButton(cursor_head, text="Добавить модель…", command=self._select_model_from_files).grid(row=0, column=2, padx=3)
        self.cursor_chat_display = tk.Text(right, wrap="word", state="disabled", font=("Segoe UI", 10), padx=12, pady=12, background="#ffffff", foreground="#202123", borderwidth=0, highlightthickness=1, highlightbackground="#e5e7eb")
        self.cursor_chat_display.grid(row=1, column=0, sticky="nsew", pady=(8, 4))
        cursor_scroll = ttk.Scrollbar(right, orient="vertical", command=self.cursor_chat_display.yview); cursor_scroll.grid(row=1, column=1, sticky="ns", pady=(8, 4)); self.cursor_chat_display.configure(yscrollcommand=cursor_scroll.set)
        self.cursor_chat_display.tag_configure("user", foreground="#08765b", spacing1=10, spacing3=6, font=("Segoe UI Semibold", 10, "bold")); self.cursor_chat_display.tag_configure("assistant", foreground="#202123", spacing1=8, spacing3=14); self.cursor_chat_display.tag_configure("system", foreground="#6b7280", spacing1=8, spacing3=8)
        self.cursor_attached_label = tk.StringVar(value="Файл не прикреплён")
        ttk.Label(right, textvariable=self.cursor_attached_label).grid(row=2, column=0, sticky="w")
        cursor_attach = ttk.Frame(right); cursor_attach.grid(row=3, column=0, sticky="ew")
        RoundedButton(cursor_attach, text="Прикрепить…", command=self.attach_chat_file).pack(side="left", padx=(0, 4)); RoundedButton(cursor_attach, text="Очистить", command=self.clear_chat_files).pack(side="left")
        self.cursor_chat_input = tk.Text(right, height=4, wrap="word", font=("Segoe UI", 10), padx=10, pady=8, background="#ffffff", foreground="#202123", insertbackground="#202123", borderwidth=0, highlightthickness=1, highlightbackground="#d1d5db", highlightcolor="#10a37f")
        self.cursor_chat_input.grid(row=4, column=0, sticky="ew", pady=6); self.cursor_chat_input.bind("<Return>", self._chat_enter)
        cursor_controls = ttk.Frame(right); cursor_controls.grid(row=5, column=0, sticky="ew")
        self.cursor_chat_send_button = RoundedButton(cursor_controls, text="Отправить", command=self.send_chat, style="Accent.TButton"); self.cursor_chat_send_button.pack(side="left", fill="x", expand=True, padx=(0, 3))
        self.cursor_chat_edit_button = RoundedButton(cursor_controls, text="Правки файла → diff", command=self.prepare_edit); self.cursor_chat_edit_button.pack(side="left", padx=3)
        self.cursor_chat_save_button = RoundedButton(cursor_controls, text="Сохранить ответ", command=self.save_chat_answer); self.cursor_chat_save_button.pack(side="left", padx=3)
        split.add(left, weight=1); split.add(center, weight=3); split.add(right, weight=2)
        self.cursor_chat_display.tag_configure("user", foreground="#08765b", spacing1=10, spacing3=6, font=("Segoe UI Semibold", 10, "bold")); self.cursor_chat_display.tag_configure("assistant", foreground="#202123", spacing1=8, spacing3=14); self.cursor_chat_display.tag_configure("system", foreground="#6b7280", spacing1=8, spacing3=8)
        self._refresh_cursor_projects(); self._refresh_cursor_chats()
    def _main_tab_changed(self, _event=None):
        selected_tab = self.main_notebook.select()
        cursor_active = selected_tab == str(self.cursor_tab)
        if cursor_active:
            self.chat_display, self.chat_input = self.cursor_chat_display, self.cursor_chat_input
            self.chat_send_button, self.chat_edit_button, self.chat_save_button = self.cursor_chat_send_button, self.cursor_chat_edit_button, self.cursor_chat_save_button
            self.chat_folder_label = self.cursor_folder_label; self.attached_label = self.cursor_attached_label
            self._cursor_editor_tab_changed()
            self._refresh_cursor_chats(); self._render_chat()
        elif selected_tab == str(self.files_tab):
            self.chat_display, self.chat_input = self.files_chat_display, self.files_chat_input
            self.chat_send_button, self.chat_edit_button, self.chat_save_button = self.files_chat_send_button, self.files_chat_edit_button, self.files_chat_save_button
            self.chat_folder_label, self.attached_label = self.files_chat_folder_label, self.files_attached_label
            self._refresh_files_chat_selector(); self._update_attachment_label(); self._render_chat()
        else:
            self.chat_display, self.chat_input = self.main_chat_display, self.main_chat_input
            self.chat_send_button, self.chat_edit_button, self.chat_save_button = self.main_chat_send_button, self.main_chat_edit_button, self.main_chat_save_button
            self.chat_folder_label, self.attached_label = self.main_chat_folder_label, self.main_attached_label
            self._render_chat()
        self._sync_chat_controls()
    def _sync_chat_controls(self):
        llm_busy = self.chat_busy or getattr(self, "cursor_completion_busy", False) or getattr(self, "legacy_ai_busy", False)
        enabled = not llm_busy
        for name in ("main_chat_send_button", "cursor_chat_send_button", "files_chat_send_button",
                     "main_chat_edit_button", "cursor_chat_edit_button", "files_chat_edit_button"):
            button = getattr(self, name, None)
            if button is not None:
                button.configure(state="normal" if enabled else "disabled")
        for name in ("main_chat_save_button", "cursor_chat_save_button", "files_chat_save_button"):
            button = getattr(self, name, None)
            if button is not None:
                button.configure(state="normal" if self.last_assistant_answer else "disabled")
        progress = getattr(self, "llm_progress", None)
        if progress is not None:
            if llm_busy:
                progress.grid(); progress.start(12)
            else:
                progress.stop(); progress.grid_remove()
    def _refresh_files_chat_selector(self):
        if not hasattr(self, "files_chat_box"): return
        ordered = sorted(self.chats, key=lambda chat: chat.get("updated", 0), reverse=True)
        self.files_chat_ids = [chat["id"] for chat in ordered]
        self.files_chat_box.configure(values=tuple(chat.get("title") or "Новый чат" for chat in ordered))
        active = next((index for index, chat_id in enumerate(self.files_chat_ids) if chat_id == self.active_chat_id), -1)
        if active >= 0: self.files_chat_box.current(active)
    def _select_files_chat(self, _event=None):
        if self.chat_busy: return
        index = self.files_chat_box.current()
        if 0 <= index < len(self.files_chat_ids): self._activate_chat(self.files_chat_ids[index])
    def _refresh_cursor_projects(self):
        if not hasattr(self, "cursor_project_box"): return
        projects = list(dict.fromkeys(p for p in self.chat_project_dirs if Path(p).is_dir()))
        self.cursor_projects = projects; self.cursor_project_box.configure(values=tuple(projects))
        self.cursor_folder_label.set(self.chat_working_dir if self.chat_working_dir in projects else (projects[-1] if projects else "Выберите рабочую папку проекта"))
    def _clear_cursor_search_results(self):
        self.cursor_search_generation = getattr(self, "cursor_search_generation", 0) + 1
        if hasattr(self, "cursor_search_results"):
            self.cursor_search_results.delete(*self.cursor_search_results.get_children()); self.cursor_search_results.grid_remove(); self.cursor_search_title.grid_remove()
            self.cursor_search_hits = {}
    def _refresh_cursor_chats(self):
        if not hasattr(self, "cursor_chat_box"): return
        self.cursor_chat_ids = [c["id"] for c in sorted(self.chats, key=lambda c: c.get("updated", 0), reverse=True)]
        self.cursor_chat_box.configure(values=tuple((c.get("title") or "Новый чат") for c in sorted(self.chats, key=lambda c: c.get("updated", 0), reverse=True)))
        active = next((i for i, cid in enumerate(self.cursor_chat_ids) if cid == self.active_chat_id), -1)
        if active >= 0: self.cursor_chat_box.current(active)
    def _select_cursor_chat(self, _event=None):
        if self.chat_busy: return
        index = self.cursor_chat_box.current()
        if 0 <= index < len(self.cursor_chat_ids): self._activate_chat(self.cursor_chat_ids[index])
    def _select_cursor_project(self, _event=None):
        path = self.cursor_folder_label.get()
        if path and Path(path).is_dir():
            self.chat_working_dir = str(Path(path).resolve()); self.project_file_path = None
            self.chat_folder_label.set(self.chat_working_dir); self.config_data["chat_working_dir"] = self.chat_working_dir; write_config(self.config_data)
            self.refresh_cursor_files(); self.refresh_project_files(); self._save_current_chat()
    def add_cursor_project(self):
        folder = filedialog.askdirectory(title="Добавить проект", initialdir=self.chat_working_dir if Path(self.chat_working_dir).is_dir() else str(Path.home()))
        if not folder: return
        folder = str(Path(folder).resolve())
        projects = list(dict.fromkeys([*self.chat_project_dirs, folder]))
        self.chat_project_dirs = projects
        self.chat_working_dir = folder; self.project_file_path = None
        self.chat_folder_label.set(folder); self.cursor_folder_label.set(folder)
        self.config_data["chat_working_dir"] = folder; write_config(self.config_data)
        self._refresh_cursor_projects(); self._clear_cursor_search_results(); self.refresh_cursor_files(); self.refresh_project_files(); self._save_current_chat()
    def new_cursor_file(self):
        roots = [p for p in self.chat_project_dirs if Path(p).is_dir()]
        if not roots:
            messagebox.showinfo(APP_NAME, "Сначала добавьте папку проекта."); return
        path = filedialog.asksaveasfilename(title="Новый файл проекта", initialdir=self.chat_working_dir if Path(self.chat_working_dir).is_dir() else roots[0],
                                            initialfile="new_file.py", filetypes=[("Все файлы", "*.*")])
        if not path: return
        target = Path(path).resolve()
        if not any(target.is_relative_to(Path(root).resolve()) for root in roots):
            messagebox.showwarning(APP_NAME, "Новый файл можно создать только внутри одного из открытых проектов."); return
        try:
            with target.open("x", encoding="utf-8", newline="") as stream: stream.write("")
        except FileExistsError:
            self.open_cursor_file(target); return
        except OSError as e:
            messagebox.showerror(APP_NAME, f"Не удалось создать файл: {e}"); return
        self.refresh_cursor_files(); self.open_cursor_file(target)
    def _cursor_editor_record(self, path=None):
        key = str(Path(path or self.cursor_active_path).resolve()) if (path or self.cursor_active_path) else None
        return self.cursor_open_files.get(key) if key else None
    def open_cursor_file(self, path, line=None):
        path = Path(path).resolve()
        if not path.is_file(): return
        roots = [Path(root).resolve() for root in self.chat_project_dirs if Path(root).is_dir()]
        if roots and not any(path.is_relative_to(root) for root in roots):
            messagebox.showwarning(APP_NAME, "Редактор открывает файлы только из подключённых проектов."); return
        key = str(path)
        record = self.cursor_open_files.get(key)
        if record:
            self.cursor_editor_tabs.select(record["frame"])
            self.cursor_active_path = key
            if line: self._cursor_goto_line(record, line)
            self.project_file_path = path
            self._cursor_editor_tab_changed()
            return
        try:
            if path.stat().st_size > 5_000_000: raise ValueError("Редактор открывает текстовые файлы размером до 5 МБ.")
            if is_plain_text_file(path):
                text, encoding = read_editable_text(path); editable = True
            elif path.suffix.lower() in SUPPORTED:
                text, encoding, editable = extract_text(path)[:100000], "utf-8", False
            else:
                text, encoding, editable = "Для этого формата доступен чат по файлу, но нет текстового редактора.", "utf-8", False
        except (OSError, UnicodeError, ValueError, ImportError, zipfile.BadZipFile) as e:
            messagebox.showerror(APP_NAME, f"Не удалось открыть файл: {e}"); return
        newline = "\r\n" if "\r\n" in text else "\n"
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        frame = ttk.Frame(self.cursor_editor_tabs); frame.columnconfigure(1, weight=1); frame.rowconfigure(0, weight=1)
        gutter = tk.Text(frame, width=5, wrap="none", state="disabled", takefocus=False, font=("Cascadia Code", 10), background="#f7f7f8", foreground="#8b9098", borderwidth=0, padx=8, pady=10)
        gutter.grid(row=0, column=0, sticky="ns")
        editor = tk.Text(frame, wrap="none", undo=True, maxundo=-1, autoseparators=True, font=("Cascadia Code", 10), background="#ffffff", foreground="#263238", insertbackground="#10a37f", selectbackground="#dff3ed", borderwidth=0, padx=12, pady=10)
        editor.grid(row=0, column=1, sticky="nsew")
        yscroll = ttk.Scrollbar(frame, orient="vertical", command=lambda *args: (editor.yview(*args), gutter.yview(*args)))
        yscroll.grid(row=0, column=2, sticky="ns")
        xscroll = ttk.Scrollbar(frame, orient="horizontal", command=editor.xview); xscroll.grid(row=1, column=1, sticky="ew")
        editor.configure(yscrollcommand=lambda first, last: (yscroll.set(first, last), gutter.yview_moveto(first)), xscrollcommand=xscroll.set)
        for tag, color in (("keyword", "#7c3aed"), ("string", "#15803d"), ("comment", "#87909b"), ("number", "#c2410c"), ("variable", "#0f766e"), ("builtin", "#2563eb"), ("tag", "#b45309"), ("blade", "#be185d")):
            editor.tag_configure(tag, foreground=color)
        if editable: editor.insert("1.0", text); editor.edit_modified(False)
        else: editor.insert("1.0", text); editor.configure(state="disabled")
        record = {"path": path, "frame": frame, "editor": editor, "gutter": gutter, "encoding": encoding,
                  "newline": newline, "saved_text": text, "editable": editable, "highlight_after": None, "signature": self._cursor_file_signature(path)}
        self.cursor_open_files[key] = record
        self.cursor_editor_tabs.add(frame, text=path.name + ("" if editable else " (просмотр)"))
        self.cursor_editor_tabs.select(frame); self.cursor_active_path = key; self.project_file_path = path
        if editable:
            editor.bind("<<Modified>>", lambda _e, k=key: self._cursor_editor_changed(k))
            editor.bind("<Control-s>", lambda _e: (self.save_cursor_file(), "break")[1])
            editor.bind("<Control-f>", lambda _e: (self.find_in_cursor_editor(), "break")[1])
            editor.bind("<Control-h>", lambda _e: (self.replace_in_cursor_editor(), "break")[1])
            editor.bind("<Control-g>", lambda _e: (self.goto_cursor_line(), "break")[1])
            editor.bind("<Control-a>", lambda _e: (editor.tag_add("sel", "1.0", "end-1c"), "break")[1])
            editor.bind("<Control-w>", lambda _e: (self.close_cursor_editor(), "break")[1])
            editor.bind("<Return>", lambda event, k=key: self._cursor_auto_indent(event, k))
            editor.bind("<Control-Button-1>", lambda event, k=key: self._cursor_go_to_definition(event, k))
            editor.bind("<F12>", lambda _e, k=key: (self._cursor_go_to_definition(None, k), "break")[1])
            editor.bind("<Shift-F12>", lambda _e, k=key: (self.find_cursor_usages(k), "break")[1])
            editor.bind("<Button-3>", lambda event, k=key: self._cursor_symbol_context_menu(event, k))
            editor.bind("<KeyRelease>", lambda _e, k=key: self._update_cursor_line_numbers(self.cursor_open_files.get(k)))
        self._update_cursor_line_numbers(record); self._cursor_editor_tab_changed(); self._highlight_cursor_editor(key)
        if line: self._cursor_goto_line(record, line)
        if self.main_notebook.select() == str(self.cursor_tab): editor.focus_set()
    def _cursor_auto_indent(self, event, key):
        record = self.cursor_open_files.get(key)
        if not record or not record["editable"]: return
        editor = record["editor"]
        line = editor.get("insert linestart", "insert")
        indent = re.match(r"[ \t]*", line).group(0)
        trimmed = line.strip()
        if record["path"].suffix.lower() == ".php" and re.match(r"(?:\}|\]|\)|else\b|elseif\b|catch\b|finally\b|end(?:if|foreach|for|while|switch|section|push|auth)\b)", trimmed, re.IGNORECASE):
            indent = indent[4:] if indent.endswith("    ") else indent[:-1] if indent.endswith("\t") else indent
        opens_block = line.rstrip().endswith((":", "{", "[", "("))
        if record["path"].name.lower().endswith(".blade.php") and re.match(r"@(?:if|unless|foreach|forelse|for|while|section|auth|guest|php)\b", trimmed, re.IGNORECASE): opens_block = True
        extra = "    " if opens_block else ""
        editor.insert("insert", "\n" + indent + extra)
        return "break"
    @staticmethod
    def _cursor_file_signature(path):
        try:
            stat = path.stat(); return stat.st_mtime_ns, stat.st_size
        except OSError: return None
    def _update_cursor_line_numbers(self, record):
        if not record: return
        editor, gutter = record["editor"], record["gutter"]
        line_count = int(editor.index("end-1c").split(".")[0]); numbers = "\n".join(str(i) for i in range(1, max(1, line_count) + 1))
        gutter.configure(state="normal"); gutter.delete("1.0", "end"); gutter.insert("1.0", numbers); gutter.configure(state="disabled")
        gutter.yview_moveto(editor.yview()[0])
    def _cursor_editor_changed(self, key):
        record = self.cursor_open_files.get(key)
        if not record: return
        editor = record["editor"]
        if editor.edit_modified():
            editor.edit_modified(False)
            dirty = editor.get("1.0", "end-1c") != record["saved_text"]
            record["dirty"] = dirty
            self.cursor_symbol_cache.pop(key, None)
            self._refresh_cursor_editor_title(record)
            self._update_cursor_line_numbers(record)
            if record.get("highlight_after"):
                try: self.after_cancel(record["highlight_after"])
                except tk.TclError: pass
            record["highlight_after"] = self.after(220, lambda k=key: self._highlight_cursor_editor(k))
    def _refresh_cursor_editor_title(self, record):
        suffix = " •" if record.get("dirty") else (" (просмотр)" if not record["editable"] else "")
        try: self.cursor_editor_tabs.tab(record["frame"], text=record["path"].name + suffix)
        except tk.TclError: pass
    def _highlight_cursor_editor(self, key):
        record = self.cursor_open_files.get(key)
        if not record or not record["editable"]: return
        editor = record["editor"]; text = editor.get("1.0", "end-1c")
        for tag in ("keyword", "string", "comment", "number", "variable", "builtin", "tag"): editor.tag_remove(tag, "1.0", "end")
        try:
            if record["path"].suffix.lower() == ".py":
                for token in tokenize.generate_tokens(io.StringIO(text).readline):
                    kind = "keyword" if token.type == tokenize.NAME and keyword.iskeyword(token.string) else {tokenize.STRING: "string", tokenize.COMMENT: "comment", tokenize.NUMBER: "number"}.get(token.type)
                    if kind: editor.tag_add(kind, f"{token.start[0]}.{token.start[1]}", f"{token.end[0]}.{token.end[1]}")
            elif record["path"].suffix.lower() == ".php":
                php_tokens = r"(?P<comment>//[^\n]*|\#(?!\[)[^\n]*|/\*[\s\S]*?\*/)|(?P<string>'(?:\\.|[^'\\])*'|\"(?:\\.|[^\"\\])*\"|`(?:\\.|[^`\\])*`)|(?P<tag><\?php|<\?=|\?>)|(?P<variable>\$[A-Za-z_]\w*)|(?P<number>\b\d+(?:\.\d+)?\b)|(?P<keyword>\b(?:abstract|and|array|as|break|callable|case|catch|class|clone|const|continue|declare|default|die|do|echo|else|elseif|empty|endfor|endforeach|endif|endswitch|endwhile|enum|eval|exit|extends|final|finally|fn|for|foreach|function|global|goto|if|implements|include|include_once|instanceof|insteadof|interface|isset|list|match|namespace|new|or|print|private|protected|public|readonly|require|require_once|return|static|switch|throw|trait|try|unset|use|var|while|xor|yield|true|false|null)\b)|(?P<builtin>\b(?:count|strlen|str_contains|str_replace|array_map|array_filter|json_encode|json_decode|htmlspecialchars|preg_match|preg_replace|file_get_contents|file_put_contents|DateTime|PDO|Exception)\b)"
                for match in re.finditer(php_tokens, text):
                    kind = match.lastgroup
                    if kind: editor.tag_add(kind, f"1.0+{match.start()}c", f"1.0+{match.end()}c")
                if record["path"].name.lower().endswith(".blade.php"):
                    for match in re.finditer(r"@(?:extends|section|endsection|yield|include|if|elseif|else|endif|foreach|endforeach|forelse|empty|endforelse|csrf|method|auth|guest|endauth|push|endpush|stack|php|endphp|vite)\b|\{\{[\s\S]*?\}\}|\{!![\s\S]*?!!\}", text):
                        editor.tag_add("blade", f"1.0+{match.start()}c", f"1.0+{match.end()}c")
            else:
                patterns = (("comment", r"(?m)(?:#|//).*?$"), ("string", r"(?:\"(?:\\.|[^\"\\])*\"|'(?:\\.|[^'\\])*')"), ("number", r"\b\d+(?:\.\d+)?\b"), ("keyword", r"\b(?:class|def|function|return|if|else|elseif|elif|for|foreach|while|import|from|const|let|var|public|private|protected|namespace|use|echo|async|await|true|false|null|None)\b"))
                for tag, pattern in patterns:
                    for match in re.finditer(pattern, text): editor.tag_add(tag, f"1.0+{match.start()}c", f"1.0+{match.end()}c")
        except (tokenize.TokenError, IndentationError, SyntaxError): pass
    def _cursor_editor_tab_changed(self, _event=None):
        selected = self.cursor_editor_tabs.select()
        record = next((r for r in self.cursor_open_files.values() if str(r["frame"]) == selected), None)
        if not record: return
        self.cursor_active_path = str(record["path"])
        in_workspace = any(record["path"].is_relative_to(Path(root).resolve()) for root in self.chat_project_dirs if Path(root).is_dir())
        self.project_file_path = record["path"] if in_workspace else None
        self.cursor_editor_path_var.set(str(record["path"]) + (" •" if record.get("dirty") else ""))
        state = ("Изменён · Ctrl+S сохранить" if record.get("dirty") else "Только просмотр" if not record["editable"] else "Готово · Ctrl+S сохранить · Ctrl+F найти")
        if not in_workspace: state = "Файл открыт из другого проекта; он не будет передан в этот чат"
        self.cursor_editor_status.set(state + f" · {self._cursor_language(record['path'])} · {record['encoding']}")
        self._update_cursor_line_numbers(record)
    def _cursor_goto_line(self, record, line):
        editor = record["editor"]; line = max(1, min(int(line), int(editor.index("end-1c").split(".")[0])))
        editor.mark_set("insert", f"{line}.0"); editor.see(f"{line}.0"); editor.focus_set()
    def goto_cursor_line(self):
        record = self._cursor_editor_record()
        if not record or not record["editable"]: return
        line = simpledialog.askinteger("Перейти к строке", "Номер строки:", parent=self, minvalue=1)
        if line: self._cursor_goto_line(record, line)
    def _cursor_language(self, path):
        path = Path(path)
        if path.name.lower().endswith(".blade.php"): return "Laravel Blade"
        if path.suffix.lower() == ".php":
            for root_value in self.chat_project_dirs:
                root = Path(root_value).resolve()
                if not path.resolve().is_relative_to(root): continue
                current = path.parent
                while current == root or current.is_relative_to(root):
                    manifest = current / "composer.json"
                    try:
                        data = json.loads(manifest.read_text(encoding="utf-8")); requirements = {**data.get("require", {}), **data.get("require-dev", {})}
                        if "laravel/framework" in requirements: return "PHP (Laravel)"
                    except (OSError, ValueError, TypeError): pass
                    if current == root: break
                    current = current.parent
        return LANGUAGE_NAMES.get(path.suffix.lower(), "текст")
    def check_php_syntax(self):
        record = self._cursor_editor_record()
        if not record or record["path"].suffix.lower() != ".php":
            messagebox.showinfo(APP_NAME, "Откройте PHP-файл для проверки синтаксиса."); return
        if record.get("dirty"):
            messagebox.showinfo(APP_NAME, "Сначала сохраните PHP-файл (Ctrl+S), затем запустите проверку."); return
        php = shutil.which("php")
        if not php:
            self.cursor_editor_status.set("PHP CLI не найден; для редактирования и LLM-подсказок он не нужен")
            messagebox.showinfo(APP_NAME, "PHP CLI не найден в PATH. Редактор и локальные LLM-подстановки работают без него."); return
        path = record["path"]; self.cursor_editor_status.set("PHP -l проверяет синтаксис без выполнения файла…")
        def work():
            try:
                result = subprocess.run([php, "-l", str(path)], capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30, check=False)
                output = (result.stdout or result.stderr or "Проверка завершена без сообщения.").strip()
                self.after(0, lambda: self._php_lint_done(result.returncode, output))
            except (OSError, subprocess.TimeoutExpired) as e:
                self.after(0, lambda error=str(e): self._php_lint_done(1, error))
        threading.Thread(target=work, daemon=True).start()
    def _php_lint_done(self, code, output):
        self.cursor_editor_status.set("PHP: синтаксис корректен" if code == 0 else "PHP: найдена ошибка синтаксиса")
        if code == 0: messagebox.showinfo("Проверка PHP", output)
        else: messagebox.showwarning("Проверка PHP", output)
    def suggest_cursor_completion(self):
        record = self._cursor_editor_record()
        if not record or not record["editable"]:
            messagebox.showinfo(APP_NAME, "Откройте текстовый или кодовый файл для LLM-вставки."); return
        if getattr(self, "cursor_completion_busy", False) or self.chat_busy:
            self.cursor_editor_status.set("Дождитесь завершения текущего запроса"); return
        editor = record["editor"]; selection = editor.tag_ranges("sel")
        if selection: start, end = str(selection[0]), str(selection[1])
        else: start = end = editor.index("insert")
        prefix = editor.get("1.0", start).splitlines()[-35:]
        suffix = editor.get(end, "end-1c").splitlines()[:20]
        language = self._cursor_language(record["path"])
        php_note = " Для PHP соблюдай версию и стиль проекта; если версия неизвестна, избегай необязательных новых конструкций." if "PHP" in language else ""
        if language in {"PHP (Laravel)", "Laravel Blade"}:
            php_note += " Учитывай conventions Laravel, существующие routes/controllers/Eloquent-паттерны. Для Blade предпочитай экранированный вывод {{ }}; используй {!! !!} только по явной необходимости. Не придумывай зависимости."
        prompt = (f"Ты помощник автодополнения кода. Язык файла: {language}.{php_note} "
                  "Предложи только фрагмент, который нужно вставить в позицию курсора. "
                  "Не повторяй код до курсора и не оборачивай ответ в markdown. "
                  "Если выделен участок, верни текст для его замены.\n\n"
                  f"Файл: {record['path'].name}\nКод до позиции:\n{chr(10).join(prefix)[-7000:]}\n"
                  f"Выделенный текст: {editor.get(start, end) if selection else '(нет)'}\n"
                  f"Код после позиции:\n{chr(10).join(suffix)[:4000]}\n\nВставляемый фрагмент:")
        marker = "cursor_completion_" + uuid.uuid4().hex
        editor.mark_set(marker + "_start", start); editor.mark_gravity(marker + "_start", "left")
        editor.mark_set(marker + "_end", end); editor.mark_gravity(marker + "_end", "right")
        model = self.model.get().strip(); self.cursor_completion_busy = True
        self.cursor_completion_button.configure(state="disabled")
        self.cursor_editor_status.set(f"Локальная модель предлагает {language}…")
        log_event(f"CURSOR_COMPLETION_START model={model!r} file_ext={record['path'].suffix.lower()}")
        self._sync_chat_controls()
        def work():
            try:
                answer = local_chat(self.model_dir_var.get(), model, [{"role": "user", "content": prompt}])
                if not answer or not answer.strip(): raise ValueError("Модель вернула пустое дополнение")
                self.after(0, lambda: self._show_cursor_completion(record, marker, answer))
            except Exception as e:
                log_event(f"CURSOR_COMPLETION_ERROR model={model!r} error={str(e)!r}", exc_info=True)
                self.after(0, lambda error=str(e): self._cursor_completion_error(record, marker, error))
        threading.Thread(target=work, daemon=True).start()
    def _cursor_completion_error(self, record, marker, error):
        self.cursor_completion_busy = False; self.cursor_completion_button.configure(state="normal")
        self._sync_chat_controls()
        self.cursor_editor_status.set("LLM-вставка не получена")
        if record and record["editor"].winfo_exists():
            for suffix in ("_start", "_end"):
                try: record["editor"].mark_unset(marker + suffix)
                except tk.TclError: pass
        messagebox.showerror("Локальная модель недоступна", error)
    def _show_cursor_completion(self, record, marker, answer):
        self.cursor_completion_busy = False; self.cursor_completion_button.configure(state="normal")
        self._sync_chat_controls()
        code = answer.rstrip("\r\n")
        if code.lstrip().startswith("```") and code.rstrip().endswith("```"):
            fenced = code.lstrip().split("\n", 1)
            if len(fenced) == 2: code = fenced[1].rsplit("```", 1)[0].rstrip("\r\n")
        if not code.strip():
            self._cursor_completion_error(record, marker, "Модель вернула пустое дополнение"); return
        if len(code) > 20000:
            self._cursor_completion_error(record, marker, "Предложение слишком длинное для автодополнения"); return
        win = tk.Toplevel(self); win.title(f"Предложение локальной модели — {self._cursor_language(record['path'])}"); win.geometry("760x440")
        win.columnconfigure(0, weight=1); win.rowconfigure(1, weight=1)
        ttk.Label(win, text=f"Модель подготовила фрагмент для {record['path'].name}. Он будет добавлен только в несохранённый буфер редактора.", padding=10).grid(row=0, column=0, sticky="w")
        preview = tk.Text(win, wrap="none", font=("Cascadia Code", 10), background="#ffffff", foreground="#263238", borderwidth=0, padx=12, pady=10)
        preview.grid(row=1, column=0, sticky="nsew", padx=10); preview.insert("1.0", code); preview.configure(state="disabled")
        actions = ttk.Frame(win, padding=10); actions.grid(row=2, column=0, sticky="e")
        RoundedButton(actions, text="Вставить в редактор", command=lambda: self._accept_cursor_completion(record, marker, code, win), style="Accent.TButton").pack(side="left", padx=4)
        RoundedButton(actions, text="Отмена", command=lambda: self._cancel_cursor_completion(record, marker, win)).pack(side="left", padx=4)
        win.protocol("WM_DELETE_WINDOW", lambda: self._cancel_cursor_completion(record, marker, win))
    def _accept_cursor_completion(self, record, marker, code, window):
        editor = record["editor"]
        if not editor.winfo_exists() or str(record["frame"]) not in self.cursor_editor_tabs.tabs():
            window.destroy(); return
        start, end = marker + "_start", marker + "_end"
        try:
            editor.delete(start, end); editor.insert(start, code); editor.mark_set("insert", f"{start}+{len(code)}c")
            editor.see("insert"); editor.focus_set(); self.cursor_editor_tabs.select(record["frame"])
        finally:
            editor.mark_unset(start, end)
        window.destroy(); self._cursor_editor_changed(str(record["path"])); self.cursor_editor_status.set("Фрагмент вставлен в буфер · Ctrl+S для сохранения")
    def _cancel_cursor_completion(self, record, marker, window):
        try: record["editor"].mark_unset(marker + "_start", marker + "_end")
        except tk.TclError: pass
        self.cursor_completion_button.configure(state="normal"); self.cursor_completion_busy = False; window.destroy()
    def find_in_cursor_editor(self):
        record = self._cursor_editor_record()
        if not record: return
        query = simpledialog.askstring("Найти", "Текст для поиска в открытом файле:", parent=self)
        if not query: return
        editor = record["editor"]; start = editor.index("insert + 1c")
        found = editor.search(query, start, stopindex="end", nocase=True) or editor.search(query, "1.0", stopindex=start, nocase=True)
        if found:
            editor.tag_remove("find_match", "1.0", "end"); editor.tag_configure("find_match", background="#fff0a6")
            editor.tag_add("find_match", found, f"{found}+{len(query)}c"); editor.mark_set("insert", found); editor.see(found)
        else: self.cursor_editor_status.set("Совпадений не найдено")
    def replace_in_cursor_editor(self):
        record = self._cursor_editor_record()
        if not record or not record["editable"]: return
        old = simpledialog.askstring("Заменить", "Найти текст:", parent=self)
        if not old: return
        new = simpledialog.askstring("Заменить", "Заменить на:", parent=self)
        if new is None: return
        editor = record["editor"]; content = editor.get("1.0", "end-1c"); updated = content.replace(old, new)
        if updated != content:
            editor.delete("1.0", "end"); editor.insert("1.0", updated); self.cursor_editor_status.set(f"Заменено совпадений: {content.count(old)}")
        else: self.cursor_editor_status.set("Совпадений не найдено")
    def save_cursor_file(self):
        record = self._cursor_editor_record()
        if not record or not record["editable"]: return
        if not record.get("dirty"):
            self.cursor_editor_status.set("Нет несохранённых изменений"); return
        path = record["path"]; current_signature = self._cursor_file_signature(path)
        if current_signature != record.get("signature") and not messagebox.askyesno("Файл изменён снаружи", f"{path.name} изменился на диске после открытия. Перезаписать текущей версией редактора?", parent=self): return
        text = record["editor"].get("1.0", "end-1c")
        try:
            backup = path.with_name(path.name + ".backup-" + datetime.now().strftime("%Y%m%d-%H%M%S-%f"))
            shutil.copy2(path, backup)
            path.write_bytes(text.replace("\n", record.get("newline", "\n")).encode(record["encoding"]))
            record["saved_text"] = text; record["dirty"] = False; record["signature"] = self._cursor_file_signature(path)
            self._refresh_cursor_editor_title(record); self.cursor_editor_status.set(f"Сохранено · backup: {backup.name}")
            self.index.refresh_path(path); self._refresh_results()
        except (OSError, UnicodeError) as e: messagebox.showerror(APP_NAME, f"Не удалось сохранить файл: {e}", parent=self)
    def close_cursor_editor(self):
        record = self._cursor_editor_record()
        if not record: return
        if record.get("dirty"):
            decision = messagebox.askyesnocancel("Несохранённые изменения", f"Сохранить изменения в {record['path'].name} перед закрытием?", parent=self)
            if decision is None: return
            if decision:
                self.save_cursor_file()
                if record.get("dirty"): return
        self.cursor_editor_tabs.forget(record["frame"]); self.cursor_open_files.pop(str(record["path"]), None)
        self.cursor_active_path = None; self.cursor_editor_path_var.set("Откройте файл проекта")
        self.cursor_editor_status.set("Вкладка закрыта")
    def search_cursor_projects(self):
        query = self.cursor_search_var.get().strip()
        if not query: return
        self.cursor_search_generation = getattr(self, "cursor_search_generation", 0) + 1
        generation = self.cursor_search_generation; roots = list(self.chat_project_dirs)
        open_buffers = {key: record["editor"].get("1.0", "end-1c") for key, record in self.cursor_open_files.items() if record["editable"]}
        self.cursor_search_results.delete(*self.cursor_search_results.get_children()); self.cursor_search_hits = {}
        self.cursor_search_title.grid(row=2, column=0, columnspan=2, sticky="w", pady=(6, 2))
        self.cursor_search_results.grid(row=3, column=0, columnspan=2, sticky="ew")
        self.cursor_editor_status.set("Поиск по всем проектам…")
        def work():
            found = []; ignored = {".git", "__pycache__", "node_modules", ".venv", "venv", "dist", "build"}; seen = set(); bytes_read = 0
            for root in roots:
                base = Path(root)
                if not base.is_dir(): continue
                for path in base.rglob("*"):
                    try:
                        if not path.is_file() or any(part in ignored for part in path.relative_to(base).parts): continue
                        stat = path.stat()
                        if stat.st_size > 2_000_000 or bytes_read + stat.st_size > 250_000_000 or not is_plain_text_file(path): continue
                        resolved = path.resolve()
                        if not any(resolved.is_relative_to(Path(root).resolve()) for root in roots if Path(root).is_dir()): continue
                        if resolved in seen: continue
                        seen.add(resolved); bytes_read += stat.st_size
                        content = open_buffers.get(str(resolved))
                        if content is None: content, _encoding = read_editable_text(path)
                        for line_number, line in enumerate(content.splitlines(), 1):
                            if query.casefold() in line.casefold(): found.append((resolved, str(base), path.relative_to(base), line_number, line.strip()[:180]))
                            if len(found) >= 1000: break
                        if len(found) >= 1000 or bytes_read >= 250_000_000: break
                    except (OSError, UnicodeError, ValueError): continue
                if len(found) >= 1000 or bytes_read >= 250_000_000: break
            try: self.after(0, lambda: self._render_cursor_search_results(found, query, generation))
            except tk.TclError: pass
        threading.Thread(target=work, daemon=True).start()
    def _render_cursor_search_results(self, found, query, generation):
        if generation != getattr(self, "cursor_search_generation", 0): return
        self.cursor_search_results.delete(*self.cursor_search_results.get_children()); self.cursor_search_hits = {}
        for index, (path, project_name, relative, line, excerpt) in enumerate(found):
            iid = str(index); self.cursor_search_hits[iid] = (path, line)
            self.cursor_search_results.insert("", "end", iid=iid, text=f"{project_name}/{relative}", values=(line, excerpt))
        self.cursor_editor_status.set(f"Найдено строк: {len(found)} для «{query}»" if found else f"Совпадений нет: «{query}»")
    @staticmethod
    def _parse_cursor_symbols(path, content):
        suffix = Path(path).suffix.lower(); symbols = []
        patterns = [
            ("class", r"^\s*(?:(?:abstract|final|readonly|export|default|public|private|protected|static)\s+)*(?:class|trait|interface|enum|struct|record|type)\s+([A-Za-z_$][\w$]*)"),
            ("function", r"^\s*(?:(?:public|protected|private|static|abstract|final|async|export|default|func|fn|fun|def|function|inline|virtual|override|constexpr|\*|&)\s+)*(?:function\s*&?|func\s+(?:\([^)]*\)\s*)?|(?:def|fn|fun)\s+)([A-Za-z_$][\w$]*)\s*(?:<[^>]*>\s*)?\("),
            ("function", r"^\s*(?:(?:export|default|const|let|var|public|private|static|async)\s+)+(?:const\s+|let\s+|var\s+)?([A-Za-z_$][\w$]*)\s*=\s*(?:async\s*)?(?:\([^)]*\)|[A-Za-z_$][\w$]*)\s*=>"),
            ("function", r"^\s*(?:(?:public|protected|private|static|abstract|final|async|export|virtual|override|inline|extern|constexpr|friend|\*|&)\s+)*(?:[A-Za-z_$][\w.$<>\[\],?]*\s+)+([A-Za-z_$][\w$]*)\s*\([^;]*\)\s*(?:const\s*)?(?:throws\s+[^{}]+)?(?:\{|$)"),
        ]
        if suffix in {".php", ".py", ".js", ".jsx", ".ts", ".tsx", ".vue", ".svelte", ".java", ".cs", ".go", ".rs", ".rb", ".swift", ".kt", ".c", ".h", ".cpp", ".hpp"}:
            for line_number, line in enumerate(content.splitlines(), 1):
                stripped = line.lstrip()
                if not stripped or stripped.startswith(("//", "#", "*", "/*")): continue
                for kind, pattern in patterns:
                    match = re.search(pattern, line)
                    if match:
                        name = match.group(1)
                        if name not in {"if", "for", "while", "switch", "catch", "return", "new"}:
                            symbols.append((line_number, kind, name))
                        break
        return symbols
    def go_to_symbol(self):
        if getattr(self, "cursor_symbol_win", None) and self.cursor_symbol_win.winfo_exists():
            self.cursor_symbol_win.deiconify(); self.cursor_symbol_query.set(""); self.cursor_symbol_tree.focus_set(); return
        roots = [Path(root).resolve() for root in self.chat_project_dirs if Path(root).is_dir()]
        if not roots:
            messagebox.showinfo(APP_NAME, "Сначала добавьте папки проектов."); return
        win = tk.Toplevel(self); win.title("Перейти к классу или функции"); win.geometry("850x540"); win.columnconfigure(0, weight=1); win.rowconfigure(2, weight=1)
        ttk.Label(win, text="Ищет классы, интерфейсы и функции по всем файлам подключённых проектов.", padding=(10, 8)).grid(row=0, column=0, sticky="w")
        query_var = tk.StringVar()
        query = RoundedEntry(win, textvariable=query_var); query.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 7)); query.focus_set()
        tree = ttk.Treeview(win, columns=("location",), show="tree headings", selectmode="browse")
        tree.heading("#0", text="Символ"); tree.heading("location", text="Расположение"); tree.column("#0", width=240); tree.column("location", width=560)
        tree.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        scroll = ttk.Scrollbar(win, orient="vertical", command=tree.yview); scroll.grid(row=2, column=1, sticky="ns", pady=(0, 10)); tree.configure(yscrollcommand=scroll.set)
        self.cursor_symbol_win = win; self.cursor_symbol_tree = tree; self.cursor_symbol_query = query_var; self.cursor_symbol_rows = []
        label = ttk.Label(win, text="Индексация символов локально…", style="Muted.TLabel", padding=(10, 0, 10, 8)); label.grid(row=3, column=0, sticky="w")
        open_buffers = {key: record["editor"].get("1.0", "end-1c") for key, record in self.cursor_open_files.items() if record["editable"]}
        old_cache = dict(self.cursor_symbol_cache); supported = {".php", ".py", ".js", ".jsx", ".ts", ".tsx", ".vue", ".svelte", ".java", ".cs", ".go", ".rs", ".rb", ".swift", ".kt", ".c", ".h", ".cpp", ".hpp"}
        def work():
            rows = []; cache = {}; ignored = {".git", "node_modules", ".venv", "venv", "vendor", "dist", "build", "__pycache__"}; scanned = 0
            for base in roots:
                for folder, dirs, files in os.walk(base):
                    dirs[:] = sorted(d for d in dirs if d not in ignored)
                    for name in files:
                        path = Path(folder) / name
                        try:
                            if path.suffix.lower() not in supported or path.stat().st_size > 2_000_000: continue
                            resolved = path.resolve()
                            if not resolved.is_relative_to(base): continue
                            signature = self._cursor_file_signature(resolved); key = str(resolved)
                            if key in open_buffers:
                                content = open_buffers[key]; symbols = self._parse_cursor_symbols(resolved, content)
                            elif key in old_cache and old_cache[key][0] == signature:
                                symbols = old_cache[key][1]
                            else:
                                if not is_plain_text_file(resolved): continue
                                content, _encoding = read_editable_text(resolved); symbols = self._parse_cursor_symbols(resolved, content)
                            cache[key] = (signature, symbols); scanned += 1
                            for line_number, kind, symbol in symbols: rows.append((symbol.casefold(), symbol, kind, resolved, line_number, str(base)))
                            if scanned >= 20000: break
                        except (OSError, UnicodeError, ValueError): continue
                    if scanned >= 20000: break
                if scanned >= 20000: break
            try: self.after(0, lambda: self._show_cursor_symbols(win, label, tree, query_var, rows, cache))
            except tk.TclError: pass
        threading.Thread(target=work, daemon=True).start()
        query_var.trace_add("write", lambda *_: self._filter_cursor_symbols())
        tree.bind("<Double-1>", lambda _e: self._open_cursor_symbol(win))
        tree.bind("<Return>", lambda _e: self._open_cursor_symbol(win))
    def _show_cursor_symbols(self, window, label, tree, query_var, rows, cache):
        if not window.winfo_exists(): return
        self.cursor_symbol_cache.update(cache); self.cursor_symbol_rows = rows
        label.configure(text=f"Символов найдено: {len(rows)} · Enter или двойной щелчок для перехода")
        self._filter_cursor_symbols()
    def _filter_cursor_symbols(self):
        tree = getattr(self, "cursor_symbol_tree", None)
        if not tree or not tree.winfo_exists(): return
        query = self.cursor_symbol_query.get().casefold(); tree.delete(*tree.get_children()); visible = 0
        for index, (_fold, symbol, kind, path, line, root) in enumerate(self.cursor_symbol_rows):
            if query and query not in symbol.casefold() and query not in str(path).casefold(): continue
            iid = str(index); tree.insert("", "end", iid=iid, text=f"{symbol}  ·  {kind}", values=(f"{path} : {line}",)); visible += 1
            if visible >= 2000: break
    def _open_cursor_symbol(self, window):
        tree = self.cursor_symbol_tree; selection = tree.selection()
        if not selection: return
        index = int(selection[0]); _fold, _symbol, _kind, path, line, _root = self.cursor_symbol_rows[index]
        self.open_cursor_file(path, line); window.destroy()
    @staticmethod
    def _cursor_token_at(record, event=None):
        editor = record["editor"]
        index = editor.index(f"@{event.x},{event.y}") if event is not None else editor.index("insert")
        line_number, column = (int(part) for part in index.split("."))
        line = editor.get(f"{line_number}.0", f"{line_number}.end")
        pattern = r"\$?[A-Za-z_][A-Za-z0-9_]*" if record["path"].suffix.lower() == ".php" or record["path"].name.lower().endswith(".blade.php") else r"[A-Za-z_$][A-Za-z0-9_$]*"
        for match in re.finditer(pattern, line):
            if match.start() <= column <= match.end(): return match.group(0)
        return ""
    def _cursor_symbol_context_menu(self, event, key):
        record = self.cursor_open_files.get(key)
        if not record: return "break"
        token = self._cursor_token_at(record, event)
        if not token: return "break"
        menu = tk.Menu(self, tearoff=False)
        menu.add_command(label=f"Перейти к определению «{token}»   Ctrl+клик", command=lambda: self._cursor_go_to_definition(None, key, token))
        menu.add_command(label=f"Найти использования «{token}»   Shift+F12", command=lambda: self.find_cursor_usages(key, token))
        menu.tk_popup(event.x_root, event.y_root)
        return "break"
    def _cursor_go_to_definition(self, event=None, key=None, token=None):
        record = self.cursor_open_files.get(key) if key else self._cursor_editor_record()
        if not record: return "break" if event is not None else None
        if event is not None: record["editor"].mark_set("insert", f"@{event.x},{event.y}")
        token = token or self._cursor_token_at(record, event)
        if not token:
            self.go_to_symbol(); return "break" if event is not None else None
        roots = [Path(root).resolve() for root in self.chat_project_dirs if Path(root).is_dir()]
        buffers = {name: item["editor"].get("1.0", "end-1c") for name, item in self.cursor_open_files.items() if item["editable"]}
        self.cursor_editor_status.set(f"Поиск объявления «{token}»…")
        def work():
            results = []; ignored = {".git", "node_modules", ".venv", "venv", "vendor", "dist", "build", "__pycache__"}
            supported = {".php", ".py", ".js", ".jsx", ".ts", ".tsx", ".vue", ".svelte", ".java", ".cs", ".go", ".rs", ".rb", ".swift", ".kt", ".c", ".h", ".cpp", ".hpp"}
            for root in roots:
                for folder, dirs, files in os.walk(root):
                    dirs[:] = sorted(d for d in dirs if d not in ignored)
                    for name in files:
                        path = Path(folder) / name
                        try:
                            resolved = path.resolve()
                            if not resolved.is_relative_to(root) or path.suffix.lower() not in supported or path.stat().st_size > 2_000_000: continue
                            content = buffers.get(str(resolved))
                            if content is None:
                                if not is_plain_text_file(resolved): continue
                                content, _encoding = read_editable_text(resolved)
                            for line, kind, symbol in self._parse_cursor_symbols(resolved, content):
                                same = symbol.casefold() == token.casefold() if path.suffix.lower() == ".php" else symbol == token
                                if same: results.append((resolved, line, kind))
                        except (OSError, UnicodeError, ValueError): continue
            try: self.after(0, lambda: self._cursor_definition_results(token, results))
            except tk.TclError: pass
        threading.Thread(target=work, daemon=True).start()
        return "break" if event is not None else None
    def _cursor_definition_results(self, token, results):
        if not results:
            self.cursor_editor_status.set(f"Объявление «{token}» не найдено в подключённых проектах"); return
        if len(results) == 1:
            path, line, _kind = results[0]; self.open_cursor_file(path, line); return
        win = tk.Toplevel(self); win.title(f"Определения: {token}"); win.geometry("760x380"); win.columnconfigure(0, weight=1); win.rowconfigure(0, weight=1)
        tree = ttk.Treeview(win, columns=("location",), show="tree headings"); tree.heading("#0", text="Тип"); tree.heading("location", text="Файл и строка"); tree.column("#0", width=120); tree.column("location", width=600); tree.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        for index, (path, line, kind) in enumerate(results): tree.insert("", "end", iid=str(index), text=kind, values=(f"{path} : {line}",))
        def open_selected(_event=None):
            selection = tree.selection()
            if selection:
                path, line, _kind = results[int(selection[0])]; self.open_cursor_file(path, line); win.destroy()
        tree.bind("<Double-1>", open_selected); tree.bind("<Return>", open_selected)
    def find_cursor_usages(self, key=None, token=None):
        record = self.cursor_open_files.get(key) if key else self._cursor_editor_record()
        if not record: return
        token = token or self._cursor_token_at(record)
        if not token:
            messagebox.showinfo(APP_NAME, "Поставьте курсор на имя функции, класса или переменной."); return
        roots = [Path(root).resolve() for root in self.chat_project_dirs if Path(root).is_dir()]
        if not roots:
            messagebox.showinfo(APP_NAME, "Сначала добавьте папки проектов."); return
        win = tk.Toplevel(self); win.title(f"Использования: {token}"); win.geometry("1000x560"); win.columnconfigure(0, weight=1); win.rowconfigure(1, weight=1)
        ttk.Label(win, text=f"Локальный поиск точных упоминаний «{token}» по подключённым проектам.", padding=8).grid(row=0, column=0, sticky="w")
        tree = ttk.Treeview(win, columns=("line", "context"), show="tree headings"); tree.heading("#0", text="Файл"); tree.heading("line", text="Строка"); tree.heading("context", text="Код вокруг упоминания"); tree.column("#0", width=300); tree.column("line", width=65); tree.column("context", width=580); tree.grid(row=1, column=0, sticky="nsew", padx=8, pady=6)
        scroll = ttk.Scrollbar(win, orient="vertical", command=tree.yview); scroll.grid(row=1, column=1, sticky="ns", pady=6); tree.configure(yscrollcommand=scroll.set)
        label = ttk.Label(win, text="Ищу по проектам…", style="Muted.TLabel", padding=8); label.grid(row=2, column=0, sticky="w")
        buffers = {name: item["editor"].get("1.0", "end-1c") for name, item in self.cursor_open_files.items() if item["editable"]}
        usage_rows = []
        self.cursor_usages_win, self.cursor_usages_tree, self.cursor_usages_rows = win, tree, usage_rows
        def work():
            found = []; seen = set(); ignored = {".git", "node_modules", ".venv", "venv", "vendor", "dist", "build", "__pycache__"}; bytes_read = 0
            pattern = re.compile(r"(?<![A-Za-z0-9_$])" + re.escape(token) + r"(?![A-Za-z0-9_$])", re.IGNORECASE if record["path"].suffix.lower() == ".php" else 0)
            for root in roots:
                for folder, dirs, files in os.walk(root):
                    dirs[:] = sorted(d for d in dirs if d not in ignored)
                    for name in files:
                        path = Path(folder) / name
                        try:
                            resolved = path.resolve(); stat = path.stat()
                            if not resolved.is_relative_to(root) or stat.st_size > 2_000_000 or bytes_read + stat.st_size > 250_000_000 or not is_plain_text_file(resolved) or resolved in seen: continue
                            seen.add(resolved); bytes_read += stat.st_size
                            content = buffers.get(str(resolved))
                            if content is None: content, _encoding = read_editable_text(resolved)
                            for line_number, line in enumerate(content.splitlines(), 1):
                                if pattern.search(line): found.append((resolved, line_number, line.strip()[:220]))
                                if len(found) >= 2000: break
                            if len(found) >= 2000 or bytes_read >= 250_000_000: break
                        except (OSError, UnicodeError, ValueError): continue
                    if len(found) >= 2000 or bytes_read >= 250_000_000: break
                if len(found) >= 2000 or bytes_read >= 250_000_000: break
            try: self.after(0, lambda: self._show_cursor_usages(win, label, tree, found, usage_rows))
            except tk.TclError: pass
        threading.Thread(target=work, daemon=True).start()
        def open_usage(_event=None):
            selection = tree.selection()
            if not selection: return
            index = int(selection[0]); path, line, excerpt = usage_rows[index]
            self.open_cursor_file(path, line)
            editor_record = self._cursor_editor_record(path)
            if editor_record:
                editor = editor_record["editor"]; position = "1.0"; needle = token
                found_at = editor.search(needle, f"{line}.0", stopindex=f"{line}.end", nocase=path.suffix.lower() == ".php")
                if found_at:
                    editor.tag_add("sel", found_at, f"{found_at}+{len(needle)}c"); editor.mark_set("insert", found_at); editor.see(found_at)
            win.destroy()
        tree.bind("<Double-1>", open_usage); tree.bind("<Return>", open_usage)
    def _show_cursor_usages(self, window, label, tree, found, rows):
        if not window.winfo_exists(): return
        rows.extend(found)
        for index, (path, line, excerpt) in enumerate(found): tree.insert("", "end", iid=str(index), text=str(path), values=(line, excerpt))
        label.configure(text=f"Найдено упоминаний: {len(found)}" if found else "Упоминаний не найдено")
    def _open_cursor_search_hit(self, _event=None):
        selection = self.cursor_search_results.selection()
        if not selection: return
        path, line = self.cursor_search_hits[selection[0]]; self.open_cursor_file(path, line)
    def _close_app(self):
        for record in list(self.cursor_open_files.values()):
            if not record.get("dirty"): continue
            self.cursor_editor_tabs.select(record["frame"]); self._cursor_editor_tab_changed()
            decision = messagebox.askyesnocancel("Несохранённые изменения", f"Сохранить изменения в {record['path'].name} перед выходом?", parent=self)
            if decision is None: return
            if decision:
                self.save_cursor_file()
                if record.get("dirty"): return
        self.destroy()
    def refresh_cursor_files(self):
        if not hasattr(self, "cursor_tree"): return
        self.cursor_tree.delete(*self.cursor_tree.get_children())
        ignored = {".git", "__pycache__", "node_modules", ".venv", "venv", "dist", "build"}
        inserted = set()
        for root in self.chat_project_dirs:
            base = Path(root)
            if not base.is_dir(): continue
            root_id = "root:" + str(base.resolve())
            self.cursor_tree.insert("", "end", iid=root_id, text=base.name or str(base), values=("проект",), open=True)
            count = 0
            for path in sorted(base.rglob("*"), key=lambda item: str(item).casefold()):
                try:
                    relative = path.relative_to(base)
                    if any(part in ignored for part in relative.parts): continue
                    if not path.is_file() or path.stat().st_size > 2_000_000: continue
                    resolved = str(path.resolve())
                    if not Path(resolved).is_relative_to(base.resolve()): continue
                    if resolved in inserted: continue
                    inserted.add(resolved)
                    self.cursor_tree.insert(root_id, "end", iid=resolved, text=str(relative), values=(human_size(path.stat().st_size),))
                    count += 1
                    if count >= 1500: break
                except (OSError, tk.TclError): continue
    def _preview_cursor_file(self, _event=None):
        selected = self.cursor_tree.selection()
        if selected and not selected[0].startswith("root:"):
            path = Path(selected[0]); self.project_file_path = path; self.chat_mode_var.set("Режим Cursor")
            self.open_cursor_file(path)
    def attach_cursor_file(self):
        selected = self.cursor_tree.selection()
        if not selected or selected[0].startswith("root:"):
            messagebox.showinfo(APP_NAME, "Выберите файл в дереве проекта."); return
        path = Path(selected[0]); self.project_file_path = path; self.chat_edit_path = path
        ok, reason = self._attach_path(path)
        if not ok: messagebox.showinfo(APP_NAME, f"Не удалось прикрепить файл: {reason}")
    def start_cursor_task(self):
        """Compatibility helper: Cursor requests now stay in the Cursor chat pane."""
        if not self.cursor_tree.selection(): messagebox.showinfo(APP_NAME, "Сначала выберите файл проекта."); return
        self.main_notebook.select(self.cursor_tab)
        self._main_tab_changed()
        self.cursor_chat_input.focus_set()
        self.send_chat()
    def _ensure_runtime(self):
        if self.runtime_ready:
            self.check_models()
            return
        if self.runtime_installing: return
        self.runtime_installing = True
        self.status.set("Первый запуск: устанавливаю встроенный CUDA-движок локальной модели…")
        def work():
            ok, detail = install_inference_runtime()
            def finish():
                self.runtime_installing = False
                self.runtime_ready = ok
                if ok:
                    self.status.set("Встроенный движок готов. Выберите локальную модель и отправьте запрос.")
                    self.check_models()
                else:
                    self.status.set("Не удалось установить локальный движок; подробности сохранены в журнале установки.")
                    messagebox.showerror(APP_NAME, "Не удалось установить встроенный движок моделей.\\n" + detail)
            self.after(0, finish)
        threading.Thread(target=work, daemon=True).start()
    def choose_model_dir(self):
        initial = self.model_dir_var.get() if Path(self.model_dir_var.get()).is_dir() else str(Path.home())
        folder = filedialog.askdirectory(title="Выберите папку локальных моделей", initialdir=initial)
        if not folder: return
        self.model_dir_var.set(str(Path(folder).resolve()))
        self.config_data["model_dir"] = self.model_dir_var.get(); write_config(self.config_data)
        self.check_models()
    def add_folder(self):
        folder = filedialog.askdirectory(title="Выберите папку для индексации")
        if folder and folder not in self.config_data["folders"]:
            self.config_data["folders"].append(folder); write_config(self.config_data); self._refresh_folders()
    def remove_folder(self):
        folders = self.config_data["folders"]
        if folders:
            folders.pop(); write_config(self.config_data); self._refresh_folders()
    def scan(self):
        if not self.config_data["folders"]:
            messagebox.showinfo(APP_NAME, "Сначала добавьте папку, которую нужно индексировать."); return
        self.status.set("Индексация запущена…")
        threading.Thread(target=self.index.scan, args=(list(self.config_data["folders"]), self.events.put), daemon=True).start()
    def _refresh_results(self):
        self.tree.delete(*self.tree.get_children()); self.rows = self.index.search(self.query.get())
        for i, (path, name, ext, size, mtime, body) in enumerate(self.rows):
            self.tree.insert("", "end", iid=str(i), text=name, values=(ext or "файл", human_size(size), datetime.fromtimestamp(mtime).strftime("%d.%m.%Y %H:%M")))
        count, size = self.index.counts(); self.status.set(f"В индексе: {count} файлов ({human_size(size)}). Найдено: {len(self.rows)}")
    def selected(self):
        selection = self.tree.selection()
        return self.rows[int(selection[0])] if selection else None
    def _show_selected(self, _event=None):
        row = self.selected()
        if not row: return
        self.preview.configure(state="normal"); self.preview.delete("1.0", "end")
        self.preview.insert("1.0", row[5][:20000] if row[5] else "Текст не извлечён. Откройте файл в установленной офисной программе.")
        self.preview.configure(state="disabled")
    def open_selected(self):
        row = self.selected()
        if not row: return
        try: os.startfile(row[0])
        except (AttributeError, OSError) as e: messagebox.showerror(APP_NAME, str(e))
    def rename_selected(self):
        row = self.selected()
        if not row: return
        name = tk.simpledialog.askstring("Переименовать файл", "Новое имя файла:", initialvalue=Path(row[0]).name)
        if not name: return
        target = Path(row[0]).with_name(name)
        if target.exists(): messagebox.showerror(APP_NAME, "Файл с таким именем уже существует."); return
        try: Path(row[0]).rename(target); self.index.relocate(Path(row[0]), target); self._refresh_results()
        except (OSError, UnicodeError, ValueError, zipfile.BadZipFile, ET.ParseError) as e: messagebox.showerror(APP_NAME, str(e))
    def move_selected(self):
        row = self.selected()
        if not row: return
        folder = filedialog.askdirectory(title="Куда переместить файл")
        if not folder: return
        source = Path(row[0]); target = Path(folder) / source.name
        if target.exists(): messagebox.showerror(APP_NAME, "Файл с таким именем уже есть в папке."); return
        try: shutil.move(str(source), str(target)); self.index.relocate(source, target); self._refresh_results()
        except OSError as e: messagebox.showerror(APP_NAME, str(e))
    def check_models(self):
        folder = self.model_dir_var.get().strip()
        self.config_data.update({"model_dir": folder, "model": self.model.get()}); write_config(self.config_data)
        self.status.set("Сканирую каталог локальных GGUF-моделей…")
        def work():
            catalog = discover_model_files(folder)
            self.after(0, lambda values=catalog: self._set_model_choices(values))
        threading.Thread(target=work, daemon=True).start()
    def _set_model_choices(self, names):
        self.model_catalog = dict(names)
        for item in self.config_data.get("model_files", []):
            candidate = Path(item).expanduser()
            if candidate.is_file() and candidate.suffix.lower() == ".gguf": self.model_catalog.setdefault(candidate.stem, candidate.resolve())
        choices = tuple(self.model_catalog)
        self.model_box.configure(values=choices)
        self.chat_model_box.configure(values=choices)
        if hasattr(self, "cursor_model_box"): self.cursor_model_box.configure(values=choices)
        if hasattr(self, "files_chat_model_box"): self.files_chat_model_box.configure(values=choices)
        if self.model.get() not in self.model_catalog and choices: self.model.set(choices[0])
        self.status.set(f"Найдено локальных моделей: {len(choices)}" if choices else "В выбранной папке не найдено моделей GGUF")
    def _context(self):
        selected = self.selected()
        rows = [selected] if selected else self.index.search(self.query.get(), limit=5)
        return "\n\n".join(f"Файл: {r[0]}\nТекст: {r[5][:5000]}" for r in rows if r)
    def ask_ai(self):
        question = self.ai_question.get().strip()
        if not question: return
        context = self._context()
        if not context: messagebox.showinfo(APP_NAME, "Сначала найдите документы или выберите файл."); return
        self._ai("Ответь по-русски на вопрос только по текстам найденных файлов. Если ответа в них нет, скажи об этом.\n\nВопрос: " + question + "\n\nДокументы:\n" + context)
    def organize_ai(self):
        rows = self.index.search(self.query.get(), limit=20)
        if not rows: messagebox.showinfo(APP_NAME, "Сначала найдите или проиндексируйте файлы."); return
        context = "\n".join(f"{r[0]} | {(r[5] or '')[:500]}" for r in rows)
        self._ai("Предложи порядок для этих файлов: сгруппируй по смыслу и предложи имена папок. НИЧЕГО не перемещай. Ответ по-русски, в виде конкретного плана.\n" + context)
    def _ai(self, prompt):
        model = self.model.get().strip()
        self.config_data.update({"model": model, "model_dir": self.model_dir_var.get()}); write_config(self.config_data)
        self.status.set("Локальная модель загружается и отвечает на CUDA…")
        self.legacy_ai_busy = True; self._sync_chat_controls()
        log_event(f"LEGACY_AI_REQUEST_START model={model!r} prompt_chars={len(prompt)}")
        def work():
            try:
                answer = local_chat(self.model_dir_var.get(), model, [{"role": "user", "content": prompt}]) or "Пустой ответ"
                self.after(0, lambda: self._legacy_ai_done(answer))
            except Exception as e:
                log_event(f"LEGACY_AI_REQUEST_ERROR model={model!r} error={str(e)!r}", exc_info=True)
                self.after(0, lambda error=str(e): self._legacy_ai_error(error))
            self.events.put("Готово")
        threading.Thread(target=work, daemon=True).start()
    def _legacy_ai_done(self, answer):
        self.legacy_ai_busy = False; self._sync_chat_controls(); self._answer(answer)
    def _legacy_ai_error(self, error):
        self.legacy_ai_busy = False; self._sync_chat_controls()
        self.status.set("Модель не ответила — подробности в журнале")
        messagebox.showerror("Локальная модель недоступна", f"{error}\n\nЖурнал диагностики: {app_data() / 'app.log'}")
    def open_chat(self):
        self.main_notebook.select(self.chat_tab)
        self.chat_input.focus_set()

    def attach_selected_result(self):
        row = self.selected()
        if not row:
            messagebox.showinfo(APP_NAME, "Сначала выберите найденный файл слева."); return
        ok, reason = self._attach_path(Path(row[0]))
        if not ok: messagebox.showinfo(APP_NAME, f"Не удалось прикрепить файл: {reason}")

    def _read_chats(self):
        try:
            chats = json.loads((app_data() / "chats.json").read_text(encoding="utf-8"))
            return [chat for chat in chats if isinstance(chat, dict) and chat.get("id")]
        except (OSError, ValueError, TypeError):
            return []

    def _write_chats(self):
        (app_data() / "chats.json").write_text(json.dumps(self.chats, ensure_ascii=False, indent=2), encoding="utf-8")

    def _save_current_chat(self):
        if not self.active_chat_id: return
        chat = next((item for item in self.chats if item.get("id") == self.active_chat_id), None)
        if not chat: return
        chat.update({"history": self.chat_history, "attachments": self.chat_attachments,
                     "mode": self.chat_mode_var.get(), "goal": self.goal_var.get(),
                     "working_dir": self.chat_working_dir, "project_dirs": list(self.chat_project_dirs), "last_answer": self.last_assistant_answer,
                     "updated": time.time()})
        first_user = next((item["content"].strip() for item in self.chat_history if item["role"] == "user"), "")
        if first_user and chat.get("title") in {None, "Новый чат"}:
            chat["title"] = first_user.replace("\n", " ")[:48]
        self._write_chats()
        if hasattr(self, "chat_listbox"): self._refresh_chat_list(self.active_chat_id)
        self._refresh_cursor_chats()
        self._refresh_files_chat_selector()

    def _refresh_chat_list(self, select_id=None):
        if not hasattr(self, "chat_listbox"): return
        self._updating_chat_list = True
        self.chat_listbox.delete(0, "end"); self.chat_list_ids = []
        for chat in sorted(self.chats, key=lambda item: item.get("updated", 0), reverse=True):
            self.chat_list_ids.append(chat["id"]); self.chat_listbox.insert("end", chat.get("title") or "Новый чат")
        chosen = select_id or self.active_chat_id
        if chosen in self.chat_list_ids:
            index = self.chat_list_ids.index(chosen); self.chat_listbox.selection_set(index); self.chat_listbox.see(index)
        self._updating_chat_list = False

    def _select_chat_from_list(self, _event=None):
        if getattr(self, "_updating_chat_list", False): return
        if self.chat_busy:
            self._refresh_chat_list(self.active_chat_id); return
        selection = self.chat_listbox.curselection()
        if not selection: return
        chat_id = self.chat_list_ids[selection[0]]
        if chat_id == self.active_chat_id: return
        self._save_current_chat()
        chat = next((item for item in self.chats if item.get("id") == chat_id), None)
        if not chat: return
        self._activate_chat(chat_id)

    def _activate_chat(self, chat_id):
        chat = next((item for item in self.chats if item.get("id") == chat_id), None)
        if not chat: return
        self.active_chat_id = chat_id; self.chat_history = list(chat.get("history", []))
        self._clear_cursor_search_results()
        self.chat_attachments = list(chat.get("attachments", [])); saved_mode = chat.get("mode", "Обычный чат")
        self.chat_mode_var.set("Режим Cursor" if saved_mode == "Код проекта" else saved_mode)
        self.goal_var.set(chat.get("goal", "")); self.last_assistant_answer = chat.get("last_answer", "")
        saved_dir = chat.get("working_dir", self.config_data.get("chat_working_dir", ""))
        self.chat_working_dir = saved_dir if saved_dir and Path(saved_dir).is_dir() else ""
        saved_projects = chat.get("project_dirs", [])
        self.chat_project_dirs = list(dict.fromkeys(p for p in saved_projects if isinstance(p, str) and Path(p).is_dir()))
        if self.chat_working_dir and self.chat_working_dir not in self.chat_project_dirs: self.chat_project_dirs.append(self.chat_working_dir)
        if not self.chat_project_dirs and self.chat_working_dir: self.chat_project_dirs = [self.chat_working_dir]
        self.chat_folder_label.set(self.chat_working_dir or "Рабочая папка не выбрана")
        if hasattr(self, "files_chat_folder_label"):
            self.files_chat_folder_label.set(self.chat_working_dir or "Папка не выбрана")
            self._refresh_files_chat_selector()
        if hasattr(self, "cursor_folder_label"):
            self.cursor_folder_label.set(self.chat_working_dir or "Выберите рабочую папку проекта")
            self._refresh_cursor_projects()
            self.refresh_cursor_files()
            if getattr(self, "cursor_editor_tabs", None) and self.cursor_editor_tabs.select(): self._cursor_editor_tab_changed()
        self.chat_edit_path = next((Path(item["path"]) for item in reversed(self.chat_attachments) if item.get("path") and Path(item["path"]).is_file()), None)
        self.project_file_path = self.chat_edit_path if self.chat_mode_var.get() == "Режим Cursor" else None
        self._sync_chat_controls()
        self._refresh_chat_list(chat_id); self._update_attachment_label(); self._update_chat_mode_ui(); self._render_chat()

    def _mode_changed(self, _event=None):
        if self.chat_busy:
            active = next((item for item in self.chats if item.get("id") == self.active_chat_id), None)
            self.chat_mode_var.set(active.get("mode", "Обычный чат") if active else "Обычный чат")
            return
        self._update_chat_mode_ui(); self._save_current_chat()
        if self.chat_mode_var.get() == "Работа по цели":
            goal_entry = self.files_goal_entry if self.main_notebook.select() == str(self.files_tab) and hasattr(self, "files_goal_entry") else self.goal_entry
            goal_entry.focus_set()

    def _update_chat_mode_ui(self):
        if self.chat_mode_var.get() == "Работа по цели":
            self.goal_bar.grid(row=2, column=0, sticky="ew")
            if hasattr(self, "files_goal_bar"): self.files_goal_bar.grid()
        else:
            self.goal_bar.grid_remove()
            if hasattr(self, "files_goal_bar"): self.files_goal_bar.grid_remove()
        if self.chat_mode_var.get() == "Режим Cursor":
            if not self.project_panel.winfo_manager(): self.chat_body.insert(0, self.project_panel, weight=1)
            if not self.editor_panel.winfo_manager(): self.chat_body.insert(1, self.editor_panel, weight=3)
            if self.chat_working_dir: self.refresh_project_files()
            active_path = getattr(self, "project_file_path", None)
            if active_path and active_path.is_file():
                for iid, candidate in self.project_paths.items():
                    if candidate == active_path:
                        self.project_tree.selection_set(iid); self.project_tree.see(iid); self._preview_project_file(active_path); break
        else:
            for panel in (self.project_panel, self.editor_panel):
                if panel.winfo_manager(): self.chat_body.forget(panel)
            self.project_editor.configure(state="normal"); self.project_editor.delete("1.0", "end"); self.project_editor.configure(state="disabled")
            self.project_file_label.set("Выберите файл проекта")

    def _refresh_chat_list_entry(self):
        self._refresh_chat_list(self.active_chat_id)

    def refresh_project_files(self):
        self.project_tree.delete(*self.project_tree.get_children()); self.project_paths = {}
        base = Path(self.chat_working_dir) if self.chat_working_dir else None
        if not base or not base.is_dir(): return
        ignored = {".git", ".venv", "venv", "node_modules", "__pycache__", "dist", "build", "target", ".idea"}
        index = 0
        for root, dirs, files in os.walk(base):
            dirs[:] = sorted(d for d in dirs if d not in ignored)
            for filename in sorted(files):
                path = Path(root) / filename
                try:
                    st = path.stat()
                    if st.st_size > 50_000_000: continue
                    relative = str(path.relative_to(base))
                    iid = str(index); index += 1; self.project_paths[iid] = path
                    self.project_tree.insert("", "end", iid=iid, text=relative, values=(human_size(st.st_size),))
                    if index >= 1500: return
                except OSError: continue

    def attach_project_file(self):
        selection = self.project_tree.selection()
        if not selection: messagebox.showinfo(APP_NAME, "Выберите файл проекта."); return
        path = self.project_paths.get(selection[0])
        if path:
            self.project_file_path = path; self.chat_edit_path = path
            self._preview_project_file(path)
            self._attach_path(path)

    def _preview_selected_project_file(self, _event=None):
        selection = self.project_tree.selection()
        if not selection: return
        path = self.project_paths.get(selection[0])
        if path:
            self.project_file_path = path
            self._preview_project_file(path)

    def _preview_project_file(self, path: Path):
        self.project_file_label.set(str(path))
        try:
            if path.suffix.lower() in IMAGE_EXTS:
                text = "Изображение. Прикрепите его к чату и выберите локальную vision-модель для анализа."
            elif path.suffix.lower() in {".xlsx", ".xlsm", ".xls", ".xlsb"}:
                text = excel_workbook_text(path, 100000)
            elif path.suffix.lower() == ".docx":
                text = extract_text(path)
            elif path.suffix.lower() in SUPPORTED:
                text = extract_text(path)
            elif path.suffix.lower() not in EDITABLE_TEXT and not is_plain_text_file(path):
                text = "Бинарный файл. Для этого формата нет текстового предпросмотра."
            else:
                text, _encoding = read_editable_text(path)
                text = text[:200000]
        except (OSError, UnicodeError, ImportError, ValueError, zipfile.BadZipFile) as e:
            text = f"Не удалось открыть предпросмотр: {e}"
        self.project_editor.configure(state="normal"); self.project_editor.delete("1.0", "end")
        self.project_editor.insert("1.0", text or "(Файл пуст)")
        self.project_editor.configure(state="disabled")

    def _build_chat_ui(self, win):
        self.chat_window = win
        win.columnconfigure(0, weight=1); win.rowconfigure(0, weight=1)
        outer = ttk.Panedwindow(win, orient="horizontal"); outer.grid(row=0, column=0, sticky="nsew")
        sidebar = ttk.Frame(outer, padding=8, style="Sidebar.TFrame"); sidebar.columnconfigure(0, weight=1); sidebar.rowconfigure(2, weight=1)
        main = ttk.Frame(outer); main.columnconfigure(0, weight=1); main.rowconfigure(3, weight=1)
        outer.add(sidebar, weight=1); outer.add(main, weight=4)
        ttk.Label(sidebar, text="Ваши чаты", style="Sidebar.TLabel", font=("Segoe UI Semibold", 11)).grid(row=0, column=0, sticky="w", pady=(0,6))
        RoundedButton(sidebar, text="＋ Новый чат", command=self.new_chat, style="Accent.TButton").grid(row=1, column=0, sticky="ew", pady=(0,6))
        self.chat_listbox = tk.Listbox(sidebar, exportselection=False, activestyle="dotbox", font=("Segoe UI", 10))
        self.chat_listbox.configure(background="#f0f2f5", foreground="#374151", selectbackground="#dff3ed", selectforeground="#145c49", borderwidth=0, highlightthickness=0, font=("Segoe UI", 10))
        self.chat_listbox.grid(row=2, column=0, sticky="nsew"); self.chat_listbox.bind("<<ListboxSelect>>", self._select_chat_from_list)
        chat_scroll = ttk.Scrollbar(sidebar, orient="vertical", command=self.chat_listbox.yview); chat_scroll.grid(row=2, column=1, sticky="ns")
        self.chat_listbox.configure(yscrollcommand=chat_scroll.set)
        head = ttk.Frame(main, padding=8); head.grid(row=0, column=0, sticky="ew"); head.columnconfigure(1, weight=1)
        ttk.Label(head, text="Режим:").grid(row=0, column=0)
        self.chat_mode_box = ttk.Combobox(head, textvariable=self.chat_mode_var, values=["Обычный чат", "Работа по цели", "Режим Cursor"], state="readonly", width=18)
        self.chat_mode_box.grid(row=0, column=1, sticky="w", padx=6); self.chat_mode_box.bind("<<ComboboxSelected>>", self._mode_changed)
        ttk.Label(head, text="Модель:").grid(row=0, column=2)
        self.chat_model_box = ttk.Combobox(head, textvariable=self.model, values=self.model_box.cget("values"), width=20)
        self.chat_model_box.grid(row=0, column=3, sticky="ew", padx=6); head.columnconfigure(3, weight=1)
        self.chat_model_box.bind("<<ComboboxSelected>>", self._model_changed)
        RoundedButton(head, text="Обновить модели", command=self.check_models).grid(row=0, column=4, padx=4)
        folder_bar = ttk.Frame(main, padding=(8, 0, 8, 6)); folder_bar.grid(row=1, column=0, sticky="ew"); folder_bar.columnconfigure(1, weight=1)
        RoundedButton(folder_bar, text="Рабочая папка…", command=self.choose_chat_working_dir).grid(row=0, column=0, padx=(0,6))
        self.chat_folder_label = tk.StringVar(value=self.chat_working_dir or "Рабочая папка не выбрана")
        ttk.Label(folder_bar, textvariable=self.chat_folder_label).grid(row=0, column=1, sticky="ew")
        RoundedButton(folder_bar, text="Открыть папку", command=self.open_chat_working_dir).grid(row=0, column=2, padx=6)
        self.goal_bar = ttk.Frame(main, padding=(8, 0, 8, 6)); self.goal_bar.columnconfigure(1, weight=1)
        ttk.Label(self.goal_bar, text="Цель:").grid(row=0, column=0, sticky="w")
        self.goal_entry = RoundedEntry(self.goal_bar, textvariable=self.goal_var)
        self.goal_entry.grid(row=0, column=1, sticky="ew", padx=6); self.goal_entry.bind("<FocusOut>", lambda _e: self._save_current_chat()); self.goal_entry.bind("<Return>", lambda _e: self._save_current_chat())
        body = ttk.Panedwindow(main, orient="horizontal"); body.grid(row=3, column=0, sticky="nsew", padx=8)
        self.project_panel = ttk.Frame(body, padding=6); self.project_panel.columnconfigure(0, weight=1); self.project_panel.rowconfigure(1, weight=1)
        ttk.Label(self.project_panel, text="Файлы проекта").grid(row=0, column=0, sticky="w")
        self.project_tree = ttk.Treeview(self.project_panel, columns=("size",), show="tree headings", selectmode="browse")
        self.project_tree.heading("#0", text="Путь"); self.project_tree.heading("size", text="Размер")
        self.project_tree.column("#0", width=230); self.project_tree.column("size", width=70)
        self.project_tree.grid(row=1, column=0, sticky="nsew"); self.project_tree.bind("<<TreeviewSelect>>", self._preview_selected_project_file); self.project_tree.bind("<Double-1>", lambda _e: self.attach_project_file())
        project_scroll = ttk.Scrollbar(self.project_panel, orient="vertical", command=self.project_tree.yview); project_scroll.grid(row=1, column=1, sticky="ns")
        self.project_tree.configure(yscrollcommand=project_scroll.set)
        tree_actions = ttk.Frame(self.project_panel); tree_actions.grid(row=2, column=0, sticky="ew", pady=4)
        RoundedButton(tree_actions, text="Обновить", command=self.refresh_project_files).pack(side="left", padx=2)
        RoundedButton(tree_actions, text="Прикрепить выбранный", command=self.attach_project_file).pack(side="left", padx=2)
        self.editor_panel = ttk.Frame(body, padding=6); self.editor_panel.columnconfigure(0, weight=1); self.editor_panel.rowconfigure(1, weight=1)
        self.project_file_label = tk.StringVar(value="Выберите файл проекта")
        ttk.Label(self.editor_panel, textvariable=self.project_file_label, style="Title.TLabel").grid(row=0, column=0, sticky="ew", pady=(0, 5))
        self.project_editor = tk.Text(self.editor_panel, wrap="none", state="disabled", font=("Cascadia Code", 10), background="#ffffff", foreground="#374151", insertbackground="#10a37f", borderwidth=0, highlightthickness=1, highlightbackground="#e5e7eb")
        self.project_editor.grid(row=1, column=0, sticky="nsew")
        editor_scroll = ttk.Scrollbar(self.editor_panel, orient="vertical", command=self.project_editor.yview); editor_scroll.grid(row=1, column=1, sticky="ns"); self.project_editor.configure(yscrollcommand=editor_scroll.set)
        editor_actions = ttk.Frame(self.editor_panel); editor_actions.grid(row=2, column=0, sticky="ew", pady=4)
        RoundedButton(editor_actions, text="Открыть в приложении", command=self.open_project_file).pack(side="left", padx=2)
        RoundedButton(editor_actions, text="Задача с файлом → diff", command=self.prepare_edit, style="Accent.TButton").pack(side="right", padx=2)
        transcript = ttk.Frame(body); transcript.columnconfigure(0, weight=1); transcript.rowconfigure(0, weight=1)
        body.add(transcript, weight=4); self.chat_body = body
        self.chat_display = tk.Text(transcript, wrap="word", state="disabled", font=("Segoe UI", 11), padx=24, pady=18, background="#ffffff", foreground="#202123", insertbackground="#10a37f", selectbackground="#dff3ed", borderwidth=0, highlightthickness=0)
        self.chat_display.grid(row=0, column=0, sticky="nsew")
        chat_history_scroll = ttk.Scrollbar(transcript, orient="vertical", command=self.chat_display.yview)
        chat_history_scroll.grid(row=0, column=1, sticky="ns"); self.chat_display.configure(yscrollcommand=chat_history_scroll.set)
        self.chat_display.tag_configure("user", foreground="#08765b", spacing1=14, spacing3=7, font=("Segoe UI Semibold", 10, "bold"))
        self.chat_display.tag_configure("assistant", foreground="#202123", spacing1=9, spacing3=18)
        self.chat_display.tag_configure("system", foreground="#6b7280", spacing1=8, spacing3=8)
        foot = ttk.Frame(main, padding=8); foot.grid(row=4, column=0, sticky="ew"); foot.columnconfigure(0, weight=1)
        self.attached_label = tk.StringVar(value="Файл не прикреплён")
        ttk.Label(foot, textvariable=self.attached_label).grid(row=0, column=0, sticky="w")
        attachment_actions = ttk.Frame(foot); attachment_actions.grid(row=0, column=1, padx=4)
        RoundedButton(attachment_actions, text="Прикрепить файлы…", command=self.attach_chat_file).pack(side="left", padx=2)
        RoundedButton(attachment_actions, text="Очистить вложения", command=self.clear_chat_files).pack(side="left", padx=2)
        self.chat_composer = tk.Canvas(foot, height=112, background="#f7f7f8", bd=0, highlightthickness=0)
        self.chat_composer.grid(row=1, column=0, sticky="ew", pady=6)
        self.chat_input = tk.Text(self.chat_composer, height=4, wrap="word", font=("Segoe UI", 10), padx=2, pady=6, background="#ffffff", foreground="#202123", insertbackground="#202123", selectbackground="#dff3ed", borderwidth=0, relief="flat", highlightthickness=0)
        self.chat_composer_item = self.chat_composer.create_window(14, 10, anchor="nw", window=self.chat_input, width=600, height=90)
        self.chat_composer.bind("<Configure>", self._resize_chat_composer)
        self.chat_input.bind("<Return>", self._chat_enter)
        controls = ttk.Frame(foot); controls.grid(row=1, column=1, sticky="se", padx=4)
        self.chat_send_button = RoundedButton(controls, text="Отправить", command=self.send_chat, style="Accent.TButton"); self.chat_send_button.pack(fill="x", pady=2)
        self.chat_edit_button = RoundedButton(controls, text="Выполнить задачу с файлом", command=self.prepare_edit); self.chat_edit_button.pack(fill="x", pady=2)
        self.chat_save_button = RoundedButton(controls, text="Сохранить ответ как файл…", command=self.save_chat_answer, state="disabled"); self.chat_save_button.pack(fill="x", pady=2)
        ttk.Label(foot, text="Enter — отправить · Shift+Enter — новая строка · Файлы передаются локальной модели, сохранение ответа выполняется вами.").grid(row=2, column=0, columnspan=2, sticky="w")
        self._refresh_chat_list()
        if self.chats:
            latest = max(self.chats, key=lambda item: item.get("updated", 0))
            self._activate_chat(latest["id"])
        else:
            self.new_chat()
        self.chat_input.focus_set()
        self.main_chat_display, self.main_chat_input = self.chat_display, self.chat_input
        self.main_chat_send_button, self.main_chat_edit_button, self.main_chat_save_button = self.chat_send_button, self.chat_edit_button, self.chat_save_button
        self.main_chat_folder_label, self.main_attached_label = self.chat_folder_label, self.attached_label

    def _chat_enter(self, event):
        if event.state & 0x0001: return
        self.send_chat(); return "break"

    def _resize_chat_composer(self, event):
        width = max(200, event.width); height = max(92, event.height)
        self.chat_composer.delete("surface")
        radius = 17; points = []
        for cx, cy, start in ((width-radius, radius, -90), (width-radius, height-radius, 0), (radius, height-radius, 90), (radius, radius, 180)):
            for step in range(7):
                angle = math.radians(start + step * 15)
                points.extend((cx + radius * math.cos(angle), cy + radius * math.sin(angle)))
        self.chat_composer.create_polygon(*points, smooth=True, splinesteps=18, fill="#ffffff", outline="#d9dce1", width=1, tags="surface")
        self.chat_composer.tag_lower("surface")
        self.chat_composer.coords(self.chat_composer_item, 14, 10)
        self.chat_composer.itemconfigure(self.chat_composer_item, width=max(120, width - 28), height=max(60, height - 20))

    def _render_chat(self):
        if not self.chat_window or not self.chat_window.winfo_exists(): return
        for display in (getattr(self, "main_chat_display", None), getattr(self, "cursor_chat_display", None), getattr(self, "files_chat_display", None)):
            if not display or not display.winfo_exists(): continue
            display.configure(state="normal"); display.delete("1.0", "end")
            if not self.chat_history:
                mode_hint = {"Работа по цели": "Опишите цель в поле выше — чат поможет двигаться по шагам.", "Режим Cursor": "Выберите проект и файл. Запрос отправьте прямо в этой панели; правки появятся в diff и применятся только после подтверждения.", "Обычный чат": "Задавайте вопросы и прикрепляйте файлы."}.get(self.chat_mode_var.get(), "Задавайте вопросы.")
                display.insert("end", f"Локальный чат\n\n{mode_hint}\n\nМодель не имеет доступа к файловой системе и не может удалять файлы.", "system")
            for item in self.chat_history:
                label = "Вы" if item["role"] == "user" else "Локальная модель"
                display.insert("end", label + "\n", item["role"]); display.insert("end", item["content"] + "\n\n", item["role"])
            display.configure(state="disabled"); display.see("end")

    def new_chat(self):
        if self.chat_busy: return
        self._save_current_chat()
        chat = {"id": uuid.uuid4().hex, "title": "Новый чат", "history": [], "attachments": [],
                "mode": "Обычный чат", "goal": "", "working_dir": self.chat_working_dir,
                "project_dirs": list(self.chat_project_dirs), "updated": time.time()}
        self.chats.insert(0, chat); self.active_chat_id = chat["id"]
        self._activate_chat(chat["id"]); self._write_chats()
        self.chat_input.focus_set()

    def choose_chat_working_dir(self):
        if self.chat_busy: return
        initial = self.chat_working_dir if Path(self.chat_working_dir).is_dir() else str(Path.home())
        folder = filedialog.askdirectory(title="Выберите рабочую папку чата", initialdir=initial)
        if folder:
            self.chat_working_dir = str(Path(folder).resolve())
            self.project_file_path = None
            self.config_data["chat_working_dir"] = self.chat_working_dir; write_config(self.config_data)
            self.chat_folder_label.set(self.chat_working_dir)
            if hasattr(self, "cursor_folder_label"): self.cursor_folder_label.set(self.chat_working_dir)
            self._save_current_chat()
            self.refresh_cursor_files()
            if self.chat_mode_var.get() == "Режим Cursor": self.refresh_project_files()

    def open_chat_working_dir(self):
        if not self.chat_working_dir or not Path(self.chat_working_dir).is_dir():
            messagebox.showinfo(APP_NAME, "Сначала выберите рабочую папку чата."); return
        try: os.startfile(self.chat_working_dir)
        except (AttributeError, OSError) as e: messagebox.showerror(APP_NAME, str(e))

    def open_project_file(self):
        path = getattr(self, "project_file_path", None)
        if not path or not path.is_file():
            messagebox.showinfo(APP_NAME, "Сначала выберите файл проекта в дереве."); return
        if path.suffix.lower() in {".exe", ".com", ".bat", ".cmd", ".msi", ".scr"}:
            messagebox.showinfo(APP_NAME, "Запуск исполняемых файлов из предпросмотра отключён."); return
        try: os.startfile(str(path))
        except (AttributeError, OSError) as e: messagebox.showerror(APP_NAME, str(e))

    def attach_chat_file(self):
        if self.chat_busy: return
        initial = self.chat_working_dir if Path(self.chat_working_dir).is_dir() else str(Path.home())
        chosen = filedialog.askopenfilenames(title="Прикрепите файлы к чату", initialdir=initial,
            filetypes=[("Документы и изображения", "*.txt *.md *.csv *.json *.py *.docx *.xlsx *.xlsm *.xlsb *.xls *.pptx *.odt *.ods *.odp *.pdf *.png *.jpg *.jpeg *.webp *.bmp"),
                       ("Книги Excel", "*.xlsx *.xlsm *.xlsb *.xls"), ("Документы Word", "*.docx"), ("Все файлы", "*.*")])
        if not chosen: return
        skipped = []
        for raw_path in chosen:
            ok, reason = self._attach_path(Path(raw_path))
            if not ok: skipped.append(Path(raw_path).name + " (" + reason + ")")
        if skipped: messagebox.showwarning("Некоторые файлы не прикреплены", "\n".join(skipped))

    def _attach_path(self, path: Path):
        if self.chat_busy: return False, "дождитесь завершения запроса"
        path = path.resolve()
        if len(self.chat_attachments) >= 8: return False, "лимит 8 файлов"
        if any(item.get("path") == str(path) for item in self.chat_attachments):
            self.chat_edit_path = path; return True, ""
        ext = path.suffix.lower(); is_image = ext in IMAGE_EXTS
        generic_text = ext not in SUPPORTED and is_plain_text_file(path)
        if path.suffix.lower() not in SUPPORTED and not generic_text: return False, "для этого бинарного формата пока нет локального парсера"
        if self.chat_mode_var.get() == "Режим Cursor":
            roots = [Path(folder).resolve() for folder in self.chat_project_dirs if Path(folder).is_dir()]
            if not roots: return False, "сначала добавьте папку проекта"
            if not any(path.is_relative_to(root) for root in roots): return False, "в режиме кода разрешены только файлы из добавленных проектов"
        try:
            if path.stat().st_size > 50_000_000: return False, "больше 50 МБ"
            if is_image and path.stat().st_size > 20_000_000: return False, "изображение больше 20 МБ"
            text = (f"Изображение {path.name}; будет передано локальной модели с поддержкой зрения." if is_image else
                    read_editable_text(path)[0][:80000] if generic_text else
                    excel_workbook_text(path, 80000) if path.suffix.lower() in {".xlsx", ".xlsm", ".xls", ".xlsb"} else extract_text(path))
        except (OSError, UnicodeError): return False, "не удалось прочитать файл"
        if text.startswith("[Для "):
            return False, text[1:-1]
        if (not text and ext not in EDITABLE_TEXT and not generic_text) or text.startswith("[Для извлечения текста из PDF"):
            return False, "текст не извлечён"
        remaining = 80000 - sum(len(item["text"]) for item in self.chat_attachments)
        if remaining <= 0: return False, "общий лимит контекста 80 000 символов"
        attached = text[:remaining]
        self.chat_attachments.append({"name": path.name, "path": str(path), "text": attached, "image": is_image, "truncated": len(attached) < len(text)})
        self.chat_edit_path = path; self._update_attachment_label(); self._save_current_chat()
        return True, ""

    def _update_attachment_label(self):
        names = [item["name"] + (" (обрезан по лимиту)" if item.get("truncated") else "") for item in self.chat_attachments]
        label = "Вложения: " + (", ".join(names) if names else "нет")
        for widget_name in ("main_attached_label", "cursor_attached_label", "files_attached_label"):
            widget = getattr(self, widget_name, None)
            if isinstance(widget, tk.StringVar): widget.set(label)

    def clear_chat_files(self):
        if self.chat_busy: return
        self.chat_attachments = []; self.chat_edit_path = None
        self._update_attachment_label(); self._save_current_chat()

    def _chat_messages(self):
        system = "Ты помощник в локальном файловом менеджере. Отвечай на языке пользователя. У тебя нет инструментов и доступа к диску: ты не можешь перемещать, создавать или удалять файлы. Для правки верни полный текст изменённого файла по явному запросу пользователя; приложение покажет его и попросит подтверждение."
        if any(item.get("image") for item in self.chat_attachments):
            system += "\nК беседе приложено изображение. Если выбранная модель не умеет анализировать изображения, прямо сообщи об этом пользователю."
        if self.chat_mode_var.get() == "Работа по цели":
            goal = self.goal_var.get().strip()
            system += "\nРежим работы по цели: помогай последовательно достигать явно указанной цели, предлагай следующий конкретный шаг и учитывай предыдущие сообщения."
            if goal: system += "\nЦель пользователя: " + goal
        elif self.chat_mode_var.get() == "Режим Cursor":
            system += "\nРежим Cursor: анализируй только открытый в редакторе файл и явно приложенные файлы. Помогай с кодом и предлагай точные правки. Рабочая папка не означает автоматический доступ к остальным файлам проекта."
            if self.chat_project_dirs: system += "\nПроекты: " + ", ".join(Path(p).name for p in self.chat_project_dirs)
            active_path = getattr(self, "project_file_path", None)
            if self.main_notebook.select() == str(self.files_tab):
                selected_row = self.selected()
                active_path = Path(selected_row[0]) if selected_row else None
            if active_path and active_path.suffix.lower() in LANGUAGE_NAMES:
                language = self._cursor_language(active_path)
                system += f"\nЯзык открытого файла: {language}. Учитывай его синтаксис и стиль проекта."
                if "PHP" in language or language == "Laravel Blade":
                    system += " Для PHP/Laravel не меняй версии или зависимости проекта; используй существующие паттерны и API. В Blade предпочитай экранированный вывод {{ }} и избегай {!! !!}, если пользователь явно не просил обратное."
            if active_path and active_path.is_file() and not any(item.get("path") == str(active_path) for item in self.chat_attachments):
                try:
                    open_record = self._cursor_editor_record(active_path)
                    active_text = (open_record["editor"].get("1.0", "end-1c") if open_record and open_record["editable"] else
                                   "Selected image: " + active_path.name if active_path.suffix.lower() in IMAGE_EXTS else
                                   excel_workbook_text(active_path, 50000) if active_path.suffix.lower() in {".xlsx", ".xlsm", ".xls", ".xlsb"}
                                   else extract_text(active_path) if active_path.suffix.lower() in SUPPORTED
                                   else read_editable_text(active_path)[0][:50000] if is_plain_text_file(active_path)
                                   else "Binary file selected; no local content adapter is available.")
                    system += f"\n\n--- Открытый файл проекта: {active_path.name} ---\n{active_text}"
                except (OSError, UnicodeError, ValueError, ImportError, zipfile.BadZipFile): pass
        elif self.main_notebook.select() == str(self.files_tab):
            selected_row = self.selected()
            if selected_row:
                active_path = Path(selected_row[0])
                try:
                    active_text = selected_row[5] or ("К сообщению прикреплено изображение; анализируй его, если модель поддерживает зрение." if active_path.suffix.lower() in IMAGE_EXTS else excel_workbook_text(active_path, 50000) if active_path.suffix.lower() in {".xlsx", ".xlsm", ".xls", ".xlsb"} else extract_text(active_path) if active_path.suffix.lower() in SUPPORTED else read_editable_text(active_path)[0][:50000] if is_plain_text_file(active_path) else "")
                    if active_text: system += f"\n\nПользователь выбрал найденный локальный файл. Отвечай по его содержимому:\n--- {active_path.name} ---\n{active_text[:50000]}"
                except (OSError, UnicodeError, ValueError, ImportError, zipfile.BadZipFile): pass
        if self.chat_attachments:
            system += "\n\nПользователь передал содержимое файлов только для этой беседы. Рабочая папка чата: " + (Path(self.chat_working_dir).name if self.chat_working_dir else "не выбрана")
            for item in self.chat_attachments:
                system += f"\n\n--- Вложение: {item['name']} ---\n{item['text']}"
        messages = [{"role": "system", "content": system}] + self.chat_history
        image_paths = [item["path"] for item in self.chat_attachments if item.get("image") and Path(item.get("path", "")).is_file()]
        active_path = getattr(self, "project_file_path", None)
        if self.main_notebook.select() == str(self.files_tab):
            selected_row = self.selected(); active_path = Path(selected_row[0]) if selected_row else None
        if (self.chat_mode_var.get() == "Режим Cursor" or self.main_notebook.select() == str(self.files_tab)) and active_path and active_path.suffix.lower() in IMAGE_EXTS and active_path.is_file():
            if str(active_path) not in image_paths: image_paths.append(str(active_path))
        if image_paths:
            for message in reversed(messages):
                if message.get("role") == "user":
                    message["_image_paths"] = image_paths; break
        return messages

    def send_chat(self):
        if self.chat_busy: return
        prompt = self.chat_input.get("1.0", "end-1c").strip()
        if not prompt: return
        search_terms = extract_file_search_query(prompt)
        edit_words = re.search(r"\b(измени|изменить|исправь|исправить|добавь|добавить|заполни|заполнить|замени|заменить|перепиши|переписать|отредактируй|отредактировать|переведи|перевести|сократи|сократить|обнови|обновить|change|edit|update|rewrite|fill|translate|format)\b", prompt, re.IGNORECASE)
        files_tab_active = self.main_notebook.select() == str(self.files_tab)
        active_path = (getattr(self, "project_file_path", None) if self.chat_mode_var.get() == "Режим Cursor" and not files_tab_active else None)
        candidate = active_path or self.chat_edit_path
        if files_tab_active:
            row = self.selected(); candidate = Path(row[0]) if row else candidate
        elif not candidate:
            row = self.selected(); candidate = Path(row[0]) if row else None
        if search_terms is None and edit_words and candidate and candidate.is_file():
            ext = candidate.suffix.lower()
            if ext in EDITABLE_TEXT | {".docx", ".xlsx", ".xlsm", ".xls", ".xlsb"} or (ext not in SUPPORTED and is_plain_text_file(candidate)):
                self.prepare_edit(); return
        self.chat_input.delete("1.0", "end"); self.chat_history.append({"role": "user", "content": prompt}); self._render_chat()
        self._save_current_chat()
        if search_terms is not None:
            results = self.index.find_relevant(search_terms)
            if results:
                lines = [f"Нашёл файлы в проиндексированных папках ({len(results)}):"]
                for path, name, ext, size, _mtime, _body in results:
                    lines.append(f"• {name} ({ext or 'без расширения'}, {human_size(size)})\n  {path}")
                lines.append("\nВыберите нужный файл в списке слева, затем попросите чат изменить его — приложение покажет diff до сохранения.")
                answer = "\n".join(lines)
            else:
                answer = ("По этому запросу совпадений в локальном индексе нет. Поиск охватывает папки, добавленные во вкладке «Файлы и поиск» и просканированные кнопкой «Сканировать». "
                          "Добавьте нужный диск или папку и запустите сканирование, затем повторите запрос.")
            self.last_assistant_answer = answer
            self._sync_chat_controls()
            self.chat_history.append({"role": "assistant", "content": answer}); self._render_chat(); self._save_current_chat()
            self.status.set("Поиск по локальному индексу завершён")
            return
        self._run_chat(self._chat_messages(), purpose="chat")

    def save_chat_answer(self):
        if not self.last_assistant_answer:
            messagebox.showinfo(APP_NAME, "В чате ещё нет ответа для сохранения."); return
        initial = self.chat_working_dir if Path(self.chat_working_dir).is_dir() else str(Path.home())
        name = "ответ.md"
        path = filedialog.asksaveasfilename(title="Сохранить ответ модели как файл", initialdir=initial, initialfile=name,
                                            defaultextension=".md", filetypes=[("Markdown", "*.md"), ("Текст", "*.txt"), ("Все файлы", "*.*")])
        if not path: return
        target = Path(path)
        if target.exists() and not messagebox.askyesno("Подтвердить замену", f"Файл уже существует:\n{target}\n\nЗаменить его содержимое?"):
            return
        try:
            target.write_text(self.last_assistant_answer, encoding="utf-8")
            messagebox.showinfo(APP_NAME, f"Ответ сохранён:\n{target}")
        except OSError as e: messagebox.showerror(APP_NAME, f"Не удалось сохранить файл:\n{e}")

    def prepare_edit(self):
        if self.chat_busy: return
        row = self.selected()
        active_project_path = getattr(self, "project_file_path", None)
        files_tab_active = self.main_notebook.select() == str(self.files_tab)
        if not row and not (self.chat_edit_path and self.chat_edit_path.is_file()) and not (self.chat_mode_var.get() == "Режим Cursor" and active_project_path and active_project_path.is_file()):
            messagebox.showinfo(APP_NAME, "Выберите файл в списке, затем попросите изменить его."); return
        if files_tab_active and row: path = Path(row[0])
        elif self.chat_mode_var.get() == "Режим Cursor" and active_project_path and active_project_path.is_file(): path = active_project_path
        else: path = self.chat_edit_path if self.chat_edit_path and self.chat_edit_path.is_file() else Path(row[0] if row else "")
        ext = path.suffix.lower()
        excel_exts = {".xlsx", ".xlsm", ".xls", ".xlsb"}
        edit_kind = "excel" if ext in excel_exts else ("docx" if ext == ".docx" else "text")
        generic_text = ext not in SUPPORTED and is_plain_text_file(path)
        if ext not in EDITABLE_TEXT and edit_kind not in {"docx", "excel"} and not generic_text:
            messagebox.showinfo(APP_NAME, "Правки поддерживаются для текста, исходного кода, DOCX и книг Excel. Для других Office-файлов пока доступен чат по содержимому."); return
        try:
            if edit_kind == "excel":
                original = excel_workbook_text(path, 80000)
                if not original.strip(): messagebox.showinfo(APP_NAME, "В книге не найдены значения ячеек для передачи модели."); return
            elif edit_kind == "docx":
                original = docx_paragraphs(path)
                if not original: messagebox.showinfo(APP_NAME, "В документе не найдено текстовых абзацев."); return
                if len(original) > 120 or sum(len(p["text"]) for p in original) > 30000:
                    messagebox.showinfo(APP_NAME, "Для правки в прототипе выберите DOCX до 120 абзацев и 30 000 символов."); return
            else:
                editor_record = self._cursor_editor_record(path)
                if editor_record and editor_record["editable"]:
                    original = editor_record["editor"].get("1.0", "end-1c"); edit_encoding = editor_record["encoding"]
                else:
                    original, edit_encoding = read_editable_text(path)
                if len(original) > 50000: messagebox.showinfo(APP_NAME, "Для правки выберите текстовый файл не длиннее 50 000 символов."); return
        except (OSError, UnicodeError, ValueError, zipfile.BadZipFile, ET.ParseError, ImportError) as e:
            messagebox.showerror(APP_NAME, f"Не удалось прочитать файл: {e}"); return
        request = self.chat_input.get("1.0", "end-1c").strip()
        if not request:
            messagebox.showinfo(APP_NAME, "Напишите, что нужно изменить, в поле сообщения."); return
        if edit_kind == "excel":
            instruction = ("Выполни задачу с данными книги Excel. Верни только JSON вида "
                           '{"operations":[{"sheet":"Имя листа","cell":"B2","value":"новое значение"}]}. '
                           "Разрешены только точечные записи новых значений в существующие ячейки существующих листов. "
                           "Не возвращай null, не удаляй и не очищай значения, не добавляй/удаляй листы, строки или столбцы. "
                           "Сохраняй все остальные ячейки без изменений. Значение может быть текстом, числом или формулой Excel. "
                           "Если просьбу нельзя выполнить такими операциями, верни пустой список operations.\n\n"
                           f"Файл: {path.name}\nЗадача пользователя: {request}\nСодержимое книги с адресами ячеек:\n{original}")
        elif edit_kind == "docx":
            original_text = json.dumps([{"id": p["id"], "text": p["text"]} for p in original], ensure_ascii=False, indent=2)
            instruction = ("Измени только те абзацы Word, которые нужны для выполнения просьбы. Верни только JSON вида "
                           "{\"changes\":[{\"id\":1,\"text\":\"полный новый текст абзаца\"}]}. "
                           "ID бери из списка. Не перечисляй неизменённые абзацы и не добавляй markdown или пояснений.\n\n"
                           f"Файл: {path.name}\nПросьба: {request}\nАбзацы:\n{original_text}")
        else:
            instruction = ("Измени файл по просьбе пользователя. Верни ПОЛНЫЙ новый текст файла без пояснений, markdown-ограждений и сокращений. "
                           "Сохрани все части, которые не просили менять.\n\n"
                           f"Файл: {path.name}\nПросьба: {request}\nИсходный текст:\n{original}")
        self.chat_input.delete("1.0", "end")
        self.chat_history.append({"role": "user", "content": f"Выполнить задачу с файлом `{path.name}`: {request}"}); self._render_chat()
        system = "Ты выполняешь запрос только на переданных данных. У тебя нет доступа к файловой системе и инструментов удаления. Не придумывай действия вне описанной схемы."
        self._run_chat([{"role": "system", "content": system}, {"role": "user", "content": instruction}], purpose="edit", edit_path=path, original=original, edit_kind=edit_kind, edit_encoding=locals().get("edit_encoding", "utf-8"))

    def _run_chat(self, messages, purpose, edit_path=None, original=None, edit_kind="text", edit_encoding="utf-8"):
        self.chat_busy = True; self._sync_chat_controls()
        model = self.model.get().strip()
        self.config_data.update({"model": model, "model_dir": self.model_dir_var.get()}); write_config(self.config_data)
        self.status.set(f"Модель {model} загружается и отвечает на CUDA…")
        log_event(f"UI_REQUEST_QUEUED model={model!r} purpose={purpose} messages={len(messages)}")
        def work():
            try:
                answer = local_chat(self.model_dir_var.get(), model, messages)
                if not answer: raise ValueError("Модель вернула пустой ответ")
                self.after(0, lambda a=answer: self._chat_done(a, purpose, edit_path, original, edit_kind, edit_encoding))
            except Exception as e:
                log_event(f"UI_REQUEST_FAILED model={model!r} purpose={purpose} error={str(e)!r}", exc_info=True)
                self.after(0, lambda error=str(e): self._chat_error(error))
        threading.Thread(target=work, daemon=True).start()

    def _chat_done(self, answer, purpose, edit_path, original, edit_kind="text", edit_encoding="utf-8"):
        self.chat_busy = False
        self._sync_chat_controls()
        self.status.set("Ответ получен от локальной модели")
        if purpose == "edit":
            if edit_kind == "excel":
                try:
                    plan = extract_json_object(answer)
                    edited = validate_excel_operations(edit_path, plan)
                except (ValueError, TypeError, KeyError, json.JSONDecodeError, ImportError, OSError) as e:
                    self._chat_error(f"Не удалось подготовить изменения Excel: {e}"); return
            elif edit_kind == "docx":
                try:
                    plan = extract_json_object(answer)
                    by_id = {p["id"]: p for p in original}; changes = []
                    seen = set()
                    for item in plan.get("changes", []):
                        source = by_id.get(item.get("id"))
                        if not source or not isinstance(item.get("text"), str): raise ValueError("Модель вернула неизвестный абзац или неверный текст")
                        if item["id"] in seen: raise ValueError("Модель повторила ID абзаца")
                        seen.add(item["id"])
                        if len(item["text"]) > 5000: raise ValueError("Правка одного абзаца слишком длинная")
                        changes.append({**source, "new_text": item["text"]})
                    if not changes: raise ValueError("Модель не предложила ни одной правки")
                    if len(changes) > len(original): raise ValueError("Слишком много правок")
                    edited = changes
                except (ValueError, TypeError, KeyError, json.JSONDecodeError) as e:
                    self._chat_error(f"Не удалось разобрать план правок DOCX: {e}. Попробуйте ещё раз или выберите другую модель."); return
            else:
                edited = answer.strip()
                if edited.startswith("```") and edited.endswith("```"):
                    edited = edited.split("\n", 1)[1].rsplit("```", 1)[0].rstrip()
                if not edited: self._chat_error("Модель вернула пустой документ."); return
                if len(edited) > 75000: self._chat_error("Ответ модели длиннее безопасного лимита для правки (75 000 символов)."); return
            self.chat_history.append({"role": "assistant", "content": "Подготовлены изменения файла для проверки. Исходник пока не менялся."}); self._render_chat()
            self._save_current_chat()
            self._show_edit_preview(edit_path, original, edited, edit_kind, edit_encoding)
        else:
            self.last_assistant_answer = answer
            self._sync_chat_controls()
            self.chat_history.append({"role": "assistant", "content": answer}); self._render_chat()
            self._save_current_chat()

    def _chat_error(self, error):
        self.chat_busy = False
        self._sync_chat_controls()
        self.status.set("Модель не ответила — подробности в журнале")
        log_event(f"CHAT_VISIBLE_ERROR error={error!r}")
        self.chat_history.append({"role": "assistant", "content": f"Ошибка локальной модели:\n{error}\n\nДиагностика записана в журнал приложения."})
        self._render_chat(); self._save_current_chat()
        messagebox.showerror("Локальная модель недоступна", f"{error}\n\nЖурнал диагностики: {app_data() / 'app.log'}")

    def _show_edit_preview(self, path, original, edited, edit_kind, edit_encoding="utf-8"):
        win = tk.Toplevel(self); win.title(f"Проверка правок — {path.name}"); win.geometry("1000x680"); win.columnconfigure(0, weight=1); win.rowconfigure(1, weight=1)
        ttk.Label(win, text="Проверьте каждую правку. Исходник пока не менялся; резервная копия создаётся только при замене.", padding=8).grid(row=0, column=0, sticky="w")
        diff = tk.Text(win, wrap="none", font=("Consolas", 10), background="#ffffff", foreground="#202123", insertbackground="#202123", selectbackground="#dff3ed", borderwidth=0, padx=10, pady=10); diff.grid(row=1, column=0, sticky="nsew", padx=8, pady=4)
        if edit_kind == "excel":
            lines = ["Изменения ячеек (файл не менялся):"]
            for change in edited:
                lines.append(f"{change['sheet']}!{change['cell']}: {change['old']!r}  →  {change['value']!r}")
            diff.insert("1.0", "\n".join(lines))
        elif edit_kind == "docx":
            preview = []
            for change in edited:
                preview.extend(difflib.unified_diff([change["text"] + "\n"], [change["new_text"] + "\n"], fromfile=f"абзац {change['id']} — исходник", tofile=f"абзац {change['id']} — правка"))
            diff.insert("1.0", "".join(preview) or "Изменений нет.")
        else:
            diff.insert("1.0", "".join(difflib.unified_diff(original.splitlines(True), edited.splitlines(True), fromfile="исходник", tofile="версия ИИ")) or "Изменений нет.")
        diff.configure(state="disabled")
        buttons = ttk.Frame(win, padding=8); buttons.grid(row=2, column=0, sticky="e")
        RoundedButton(buttons, text="Сохранить как копию", command=lambda: self._save_edit_copy(path, edited, win, edit_kind, edit_encoding)).pack(side="left", padx=4)
        if edit_kind != "excel" or path.suffix.lower() in {".xlsx", ".xlsm"}:
            RoundedButton(buttons, text="Заменить исходник с backup", command=lambda: self._replace_with_backup(path, edited, win, edit_kind, edit_encoding)).pack(side="left", padx=4)
        RoundedButton(buttons, text="Отмена", command=win.destroy).pack(side="left", padx=4)

    def _save_edit_copy(self, path, edited, window, edit_kind, edit_encoding="utf-8"):
        target_ext = ".xlsx" if edit_kind == "excel" and path.suffix.lower() in {".xls", ".xlsb"} else path.suffix
        target = path.with_name(path.stem + "_edited" + target_ext); number = 2
        while target.exists():
            target = path.with_name(f"{path.stem}_edited_{number}{target_ext}"); number += 1
        try:
            if edit_kind == "excel": write_excel_operations(path, target, edited)
            elif edit_kind == "docx": write_docx_edits(path, target, edited)
            else: target.write_bytes(edited.encode(edit_encoding))
            self.index.refresh_path(target)
            self.chat_history.append({"role": "assistant", "content": f"Готовый файл сохранён:\n{target}"})
            self._render_chat(); self._save_current_chat()
            messagebox.showinfo(APP_NAME, f"Готовый файл сохранён:\n{target}"); window.destroy(); self._refresh_results()
        except (OSError, UnicodeError, zipfile.BadZipFile, ET.ParseError) as e: messagebox.showerror(APP_NAME, str(e))

    def _replace_with_backup(self, path, edited, window, edit_kind, edit_encoding="utf-8"):
        if not messagebox.askyesno("Подтвердить правку", f"Применить изменения к {path.name}? Перед правкой будет создана резервная копия исходного файла.\n\nМодель не может удалять файлы."):
            return
        backup = path.with_name(path.name + ".backup-" + datetime.now().strftime("%Y%m%d-%H%M%S-%f"))
        temporary = path.with_name(f".{path.name}.{uuid.uuid4().hex}.tmp")
        try:
            shutil.copy2(path, backup)
            if edit_kind == "excel": write_excel_operations(path, temporary, edited)
            elif edit_kind == "docx": write_docx_edits(path, temporary, edited)
            else:
                record = self._cursor_editor_record(path)
                newline = record.get("newline", "\n") if record else "\n"
                temporary.write_bytes(edited.replace("\n", newline).encode(edit_encoding))
            temporary.replace(path)
            editor_record = self._cursor_editor_record(path)
            if editor_record and editor_record["editable"]:
                editor = editor_record["editor"]; editor.delete("1.0", "end"); editor.insert("1.0", edited); editor.edit_modified(False)
                editor_record["saved_text"] = edited; editor_record["dirty"] = False; editor_record["signature"] = self._cursor_file_signature(path)
                self.cursor_symbol_cache.pop(str(path.resolve()), None)
                self._refresh_cursor_editor_title(editor_record); self._update_cursor_line_numbers(editor_record); self._highlight_cursor_editor(str(path.resolve()))
            self.index.refresh_path(path)
            messagebox.showinfo(APP_NAME, f"Изменения сохранены. Резервная копия:\n{backup}")
            window.destroy(); self._refresh_results()
        except (OSError, UnicodeError, ValueError, zipfile.BadZipFile, ET.ParseError) as e:
            messagebox.showerror(APP_NAME, f"Не удалось сохранить правки. Резервная копия (если создана):\n{backup}\n\n{e}")
    def _answer(self, answer):
        win = tk.Toplevel(self); win.title("Ответ локального ИИ"); win.geometry("720x520")
        text = tk.Text(win, wrap="word", padx=16, pady=16, background="#ffffff", foreground="#202123", borderwidth=0); text.pack(fill="both", expand=True); text.insert("1.0", answer); text.configure(state="disabled")
    def _poll(self):
        try:
            while True: self.status.set(self.events.get_nowait()); self._refresh_results()
        except queue.Empty: pass
        self.after(500, self._poll)


if __name__ == "__main__":
    # Needed for the standard-library simpledialog helper on all Tk builds.
    from tkinter import simpledialog
    App().mainloop()
