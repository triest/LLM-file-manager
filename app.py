"""Local-first file manager. Python 3.10+; standard-library core plus local document adapters."""
from __future__ import annotations

import json
import base64
import difflib
import os
import queue
import re
import shutil
import sqlite3
import subprocess
import sys
import threading
import time
import uuid
import urllib.error
import urllib.request
from urllib.parse import urlparse
import zipfile
from datetime import datetime
from html.parser import HTMLParser
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import tkinter as tk
import xml.etree.ElementTree as ET

APP_NAME = "Локальный файловый менеджер"
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".bmp"}
SUPPORTED = {".txt", ".md", ".rst", ".csv", ".tsv", ".json", ".xml", ".html", ".htm", ".py", ".js", ".ts", ".css", ".log", ".ini", ".yaml", ".yml", ".toml", ".rtf", ".docx", ".xlsx", ".xlsm", ".xls", ".xlsb", ".pptx", ".odt", ".ods", ".odp", ".pdf", ".png", ".jpg", ".jpeg", ".webp", ".bmp", ".sql", ".sh", ".ps1", ".bat", ".cs", ".java", ".go", ".rs", ".c", ".h", ".cpp", ".hpp", ".jsx", ".tsx", ".vue", ".svelte", ".php", ".rb", ".swift", ".kt", ".properties", ".env"}
MAX_TEXT = 500_000
OLLAMA_URL = "http://127.0.0.1:11434"
EDITABLE_TEXT = {".txt", ".md", ".rst", ".csv", ".tsv", ".json", ".xml", ".html", ".htm", ".py", ".js", ".ts", ".css", ".log", ".ini", ".yaml", ".yml", ".toml", ".sql", ".sh", ".ps1", ".bat", ".cs", ".java", ".go", ".rs", ".c", ".h", ".cpp", ".hpp", ".jsx", ".tsx", ".vue", ".svelte", ".php", ".rb", ".swift", ".kt", ".properties", ".env"}
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def app_data() -> Path:
    root = Path(os.environ.get("APPDATA", Path.home())) / "LocalFileManager"
    root.mkdir(parents=True, exist_ok=True)
    return root


def read_config() -> dict:
    p = app_data() / "config.json"
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {"folders": [], "model": "qwen2.5:7b", "provider": "Ollama", "endpoint": OLLAMA_URL}


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
    if not any(word in lowered for word in intents): return None
    local_file_context = ("комп" in lowered or "диск" in lowered or "папк" in lowered or "файл" in lowered or
                          "документ" in lowered or "где лежит" in lowered or "где находится" in lowered)
    if not local_file_context: return None
    stop = {"найди", "найти", "поищи", "отыщи", "где", "лежит", "находится", "у", "меня", "на", "в", "во", "по", "про", "для", "из", "с", "от", "моем", "моём", "компе", "компьютере", "диске", "где-то", "гдето", "пожалуйста", "мне", "его", "ее", "её", "тот", "этот", "и", "или", "моя", "моей", "мой", "мое", "моё"}
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
        terms = list(dict.fromkeys(search_fold(term) for term in terms))[:12]
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


def _local_url(endpoint: str, suffix: str) -> str:
    parsed = urlparse(endpoint.strip())
    if parsed.scheme not in {"http", "https"} or parsed.hostname not in {"127.0.0.1", "localhost", "::1"}:
        raise ValueError("Для приватности можно подключать только сервисы на этом компьютере (localhost).")
    return endpoint.rstrip("/") + suffix


def local_model_names(provider: str, endpoint: str):
    path = "/api/tags" if provider == "Ollama" else "/models"
    with urllib.request.urlopen(_local_url(endpoint, path), timeout=3) as r: payload = json.loads(r.read())
    return [m["name"] for m in payload.get("models", [])] if provider == "Ollama" else [m["id"] for m in payload.get("data", [])]


def local_chat(provider: str, endpoint: str, model: str, messages: list[dict], timeout=180):
    path = "/api/chat" if provider == "Ollama" else "/chat/completions"
    prepared = []
    for message in messages:
        item = {key: value for key, value in message.items() if key != "_image_paths"}
        image_paths = message.get("_image_paths", [])
        if image_paths:
            if provider == "Ollama":
                item["images"] = [base64.b64encode(Path(image_path).read_bytes()).decode("ascii") for image_path in image_paths]
            else:
                parts = [{"type": "text", "text": str(item.get("content", ""))}]
                for image_path in image_paths:
                    mime = "image/jpeg" if Path(image_path).suffix.lower() in {".jpg", ".jpeg"} else "image/" + Path(image_path).suffix.lower().lstrip(".")
                    encoded = base64.b64encode(Path(image_path).read_bytes()).decode("ascii")
                    parts.append({"type": "image_url", "image_url": {"url": f"data:{mime};base64,{encoded}"}})
                item["content"] = parts
        prepared.append(item)
    payload = {"model": model, "messages": prepared, "stream": False}
    req = urllib.request.Request(_local_url(endpoint, path), data=json.dumps(payload).encode(), headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as r: result = json.loads(r.read())
    return result.get("message", {}).get("content") if provider == "Ollama" else result["choices"][0]["message"]["content"]


def ollama_tags():
    return local_model_names("Ollama", OLLAMA_URL)


def try_start_ollama():
    candidates = [shutil.which("ollama"), r"C:\Program Files\Ollama\ollama.exe", str(Path.home() / "AppData/Local/Programs/Ollama/ollama.exe")]
    exe = next((c for c in candidates if c and Path(c).is_file()), None)
    if not exe: return False
    env = os.environ.copy()
    # Use the already downloaded model store found on this machine.
    if (Path("J:/OllamaModels/manifests")).is_dir(): env.setdefault("OLLAMA_MODELS", "J:/OllamaModels")
    try:
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        subprocess.Popen([exe, "serve"], env=env, creationflags=flags, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except OSError: return False


def human_size(size):
    value = float(size)
    for unit in ("Б", "КБ", "МБ", "ГБ"):
        if value < 1024 or unit == "ГБ": return f"{value:.1f} {unit}"
        value /= 1024


def search_fold(value):
    return str(value or "").casefold().replace("ё", "е")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self._configure_styles()
        self.title(APP_NAME); self.geometry("1180x760"); self.minsize(900, 600)
        self.config_data = read_config(); self.index = Index(); self.events = queue.Queue(); self.rows = []
        self.provider = tk.StringVar(value=self.config_data.get("provider", "Ollama"))
        self.endpoint = tk.StringVar(value=self.config_data.get("endpoint", OLLAMA_URL))
        self.chats = self._read_chats(); self.active_chat_id = None
        self.chat_history = []; self.chat_attachments = []; self.chat_window = None; self.chat_busy = False
        self.chat_mode_var = tk.StringVar(value="Обычный чат"); self.goal_var = tk.StringVar(value="")
        self.chat_edit_path = None
        self.project_file_path = None
        self.chat_working_dir = self.config_data.get("chat_working_dir", "")
        if self.chat_working_dir and not Path(self.chat_working_dir).is_dir(): self.chat_working_dir = ""
        self.last_assistant_answer = ""
        self._build(); self._refresh_folders(); self._refresh_results(); self._poll()
        # Start/check the local Ollama service automatically; no console is shown.
        self.after(900, self.check_models)

    def _configure_styles(self):
        style = ttk.Style(self)
        try: style.theme_use("clam")
        except tk.TclError: pass
        bg, panel, text, muted, accent = "#f3f5f9", "#ffffff", "#1f2937", "#64748b", "#4f46e5"
        self.configure(background=bg)
        style.configure(".", font=("Segoe UI", 10), background=bg, foreground=text)
        style.configure("TFrame", background=bg)
        style.configure("Surface.TFrame", background=panel)
        style.configure("Sidebar.TFrame", background="#e9edf5")
        style.configure("TLabel", background=bg, foreground=text)
        style.configure("Muted.TLabel", background=bg, foreground=muted)
        style.configure("Title.TLabel", background=bg, foreground="#111827", font=("Segoe UI Semibold", 12))
        style.configure("TButton", padding=(11, 7), background=panel, foreground=text, borderwidth=0)
        style.map("TButton", background=[("active", "#e2e8f0"), ("pressed", "#cbd5e1")])
        style.configure("Accent.TButton", padding=(13, 8), background=accent, foreground="#ffffff", font=("Segoe UI Semibold", 10))
        style.map("Accent.TButton", background=[("active", "#4338ca"), ("pressed", "#3730a3"), ("disabled", "#a5b4fc")])
        style.configure("TEntry", padding=7, fieldbackground=panel)
        style.configure("TCombobox", padding=6, fieldbackground=panel, background=panel)
        style.configure("TNotebook", background=bg, borderwidth=0)
        style.configure("TNotebook.Tab", padding=(16, 9), background="#e9edf5", foreground=muted)
        style.map("TNotebook.Tab", background=[("selected", panel)], foreground=[("selected", accent)])
        style.configure("Treeview", rowheight=27, background=panel, fieldbackground=panel, foreground=text)
        style.configure("Treeview.Heading", padding=7, background="#e9edf5", foreground=text, font=("Segoe UI Semibold", 9))

    def _build(self):
        self.columnconfigure(0, weight=1); self.rowconfigure(1, weight=1)
        bar = ttk.Frame(self, padding=10); bar.grid(row=0, column=0, sticky="ew"); bar.columnconfigure(1, weight=1)
        ttk.Label(bar, text="Папки для индексации:").grid(row=0, column=0, sticky="w")
        self.folder_var = tk.StringVar(); ttk.Entry(bar, textvariable=self.folder_var, state="readonly").grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(bar, text="Добавить папку", command=self.add_folder).grid(row=0, column=2, padx=3)
        ttk.Button(bar, text="Убрать", command=self.remove_folder).grid(row=0, column=3, padx=3)
        ttk.Button(bar, text="Сканировать", command=self.scan).grid(row=0, column=4, padx=3)
        self.main_notebook = ttk.Notebook(self); self.main_notebook.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0,10))
        self.files_tab = ttk.Frame(self.main_notebook); self.chat_tab = ttk.Frame(self.main_notebook)
        self.main_notebook.add(self.files_tab, text="Файлы и поиск")
        self.main_notebook.add(self.chat_tab, text="Чат")
        body = ttk.Panedwindow(self.files_tab, orient="horizontal"); body.grid(row=0, column=0, sticky="nsew")
        self.files_tab.columnconfigure(0, weight=1); self.files_tab.rowconfigure(0, weight=1)
        left = ttk.Frame(body, padding=6); left.columnconfigure(0, weight=1); left.rowconfigure(2, weight=1); body.add(left, weight=3)
        search = ttk.Frame(left); search.grid(row=0, column=0, sticky="ew"); search.columnconfigure(0, weight=1)
        self.query = tk.StringVar(); ent = ttk.Entry(search, textvariable=self.query); ent.grid(row=0, column=0, sticky="ew"); ent.bind("<Return>", lambda e: self._refresh_results())
        ttk.Button(search, text="Найти", command=self._refresh_results).grid(row=0, column=1, padx=4)
        ttk.Label(left, text="Ищет в именах файлов и тексте TXT, DOCX, XLSX, PPTX, ODT, CSV, PDF* и других текстовых файлов.", wraplength=500).grid(row=1, column=0, sticky="w", pady=6)
        self.tree = ttk.Treeview(left, columns=("type", "size", "date"), show="tree headings", selectmode="browse")
        self.tree.heading("#0", text="Файл"); self.tree.heading("type", text="Тип"); self.tree.heading("size", text="Размер"); self.tree.heading("date", text="Изменён")
        self.tree.column("#0", width=340); self.tree.column("type", width=75); self.tree.column("size", width=90); self.tree.column("date", width=125)
        self.tree.grid(row=2, column=0, sticky="nsew"); self.tree.bind("<<TreeviewSelect>>", self._show_selected); self.tree.bind("<Double-1>", lambda e: self.open_selected())
        actions = ttk.Frame(left); actions.grid(row=3, column=0, sticky="ew", pady=6)
        ttk.Button(actions, text="Открыть файл", command=self.open_selected).pack(side="left", padx=(0,5))
        ttk.Button(actions, text="Переименовать", command=self.rename_selected).pack(side="left", padx=5)
        ttk.Button(actions, text="Переместить…", command=self.move_selected).pack(side="left", padx=5)
        right = ttk.Frame(body, padding=8); right.columnconfigure(0, weight=1); right.rowconfigure(1, weight=1); body.add(right, weight=2)
        ttk.Label(right, text="Содержимое документа").grid(row=0, column=0, sticky="w")
        self.preview = tk.Text(right, wrap="word", font=("Segoe UI", 10), state="disabled"); self.preview.grid(row=1, column=0, sticky="nsew", pady=6)
        ttk.Label(right, text="Локальный ИИ").grid(row=2, column=0, sticky="w", pady=(6,2))
        provider_bar = ttk.Frame(right); provider_bar.grid(row=3, column=0, sticky="ew"); provider_bar.columnconfigure(1, weight=1)
        ttk.Label(provider_bar, text="Подключение").grid(row=0, column=0, sticky="w")
        self.provider_box = ttk.Combobox(provider_bar, textvariable=self.provider, values=["Ollama", "LM Studio"], state="readonly", width=12); self.provider_box.grid(row=0, column=1, sticky="ew", padx=4); self.provider_box.bind("<<ComboboxSelected>>", self._provider_changed)
        ttk.Label(right, text="Локальный адрес API").grid(row=4, column=0, sticky="w")
        ttk.Entry(right, textvariable=self.endpoint).grid(row=5, column=0, sticky="ew")
        model_bar = ttk.Frame(right); model_bar.grid(row=6, column=0, sticky="ew", pady=4); model_bar.columnconfigure(0, weight=1)
        self.model = tk.StringVar(value=self.config_data.get("model", "qwen2.5:7b"))
        choices = [self.model.get(), "qwen2.5:3b", "qwen2.5:7b", "deepseek-r1:14b", "qwen2.5vl:7b", "qwen3-vl:4b"]
        self.model_box = ttk.Combobox(model_bar, textvariable=self.model, values=list(dict.fromkeys(choices))); self.model_box.grid(row=0, column=0, sticky="ew")
        ttk.Button(model_bar, text="Найти модели", command=self.check_models).grid(row=0, column=1, padx=4)
        ttk.Button(right, text="Открыть чат", command=self.open_chat).grid(row=7, column=0, sticky="ew", pady=(4,2))
        self.ai_question = tk.StringVar(); ttk.Entry(right, textvariable=self.ai_question).grid(row=8, column=0, sticky="ew", pady=4); self.ai_question.set("Вопрос по выбранным файлам…")
        ttk.Button(right, text="Ответить по файлам", command=self.ask_ai).grid(row=9, column=0, sticky="ew")
        ttk.Button(right, text="Предложить сортировку", command=self.organize_ai).grid(row=10, column=0, sticky="ew", pady=5)
        ttk.Label(right, text="* Для PDF-поиска установите pypdf. Ollama: порт 11434. LM Studio: обычно 1234/v1.", wraplength=370).grid(row=11, column=0, sticky="w", pady=5)
        self.status = tk.StringVar(value="Добавьте папку и нажмите «Сканировать»."); ttk.Label(self, textvariable=self.status, relief="sunken", anchor="w", padding=4).grid(row=2, column=0, sticky="ew")
        self._build_chat_ui(self.chat_tab)
        self.main_notebook.select(self.chat_tab)

    def _refresh_folders(self): self.folder_var.set("; ".join(self.config_data["folders"]) or "Папки не выбраны")
    def _provider_changed(self, _event=None):
        endpoint = self.endpoint.get().strip().rstrip("/")
        if self.provider.get() == "LM Studio" and endpoint == OLLAMA_URL:
            self.endpoint.set("http://127.0.0.1:1234/v1")
        elif self.provider.get() == "Ollama" and endpoint == "http://127.0.0.1:1234/v1":
            self.endpoint.set(OLLAMA_URL)
        self.config_data.update({"provider": self.provider.get(), "endpoint": self.endpoint.get()}); write_config(self.config_data)
        return "break"
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
        provider = self.provider.get(); endpoint = self.endpoint.get().strip()
        self.config_data.update({"provider": provider, "endpoint": endpoint, "model": self.model.get()}); write_config(self.config_data)
        self.status.set(f"Поиск локальных моделей: {provider}…")
        def work():
            try:
                tags = local_model_names(provider, endpoint)
                if not tags and provider == "Ollama" and endpoint.rstrip("/") == OLLAMA_URL and try_start_ollama():
                    for _ in range(15):
                        time.sleep(1)
                        try: tags = local_model_names(provider, endpoint); break
                        except Exception: pass
                message = f"{provider}: " + (", ".join(tags) if tags else "модели не найдены")
                self.events.put(message)
                if tags: self.after(0, lambda names=tuple(tags): self._set_model_choices(names))
            except Exception as e:
                started = provider == "Ollama" and endpoint.rstrip("/") == OLLAMA_URL and try_start_ollama()
                self.events.put("Запускаю Ollama…" if started else f"Не удалось подключиться к {provider}: {e}")
        threading.Thread(target=work, daemon=True).start()
    def _set_model_choices(self, names):
        self.model_box.configure(values=names)
        if self.chat_window and self.chat_window.winfo_exists(): self.chat_model_box.configure(values=names)
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
        model = self.model.get().strip(); provider = self.provider.get(); endpoint = self.endpoint.get().strip()
        self.config_data.update({"model": model, "provider": provider, "endpoint": endpoint}); write_config(self.config_data)
        self.status.set("ИИ обрабатывает запрос локально…")
        def work():
            try:
                answer = local_chat(provider, endpoint, model, [{"role": "user", "content": prompt}]) or "Пустой ответ"
                self.after(0, lambda: self._answer(answer))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Локальная модель недоступна", f"{e}\n\nПроверьте локальный адрес, выбранную модель и её запуск."))
            self.events.put("Готово")
        threading.Thread(target=work, daemon=True).start()
    def open_chat(self):
        self.main_notebook.select(self.chat_tab)
        self.chat_input.focus_set()

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
                     "working_dir": self.chat_working_dir, "last_answer": self.last_assistant_answer,
                     "updated": time.time()})
        first_user = next((item["content"].strip() for item in self.chat_history if item["role"] == "user"), "")
        if first_user and chat.get("title") in {None, "Новый чат"}:
            chat["title"] = first_user.replace("\n", " ")[:48]
        self._write_chats()
        if hasattr(self, "chat_listbox"): self._refresh_chat_list(self.active_chat_id)

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
        self.chat_attachments = list(chat.get("attachments", [])); saved_mode = chat.get("mode", "Обычный чат")
        self.chat_mode_var.set("Режим Cursor" if saved_mode == "Код проекта" else saved_mode)
        self.goal_var.set(chat.get("goal", "")); self.last_assistant_answer = chat.get("last_answer", "")
        saved_dir = chat.get("working_dir", self.config_data.get("chat_working_dir", ""))
        self.chat_working_dir = saved_dir if saved_dir and Path(saved_dir).is_dir() else ""
        self.chat_folder_label.set(self.chat_working_dir or "Рабочая папка не выбрана")
        self.chat_edit_path = next((Path(item["path"]) for item in reversed(self.chat_attachments) if item.get("path") and Path(item["path"]).is_file()), None)
        self.project_file_path = self.chat_edit_path if self.chat_mode_var.get() == "Режим Cursor" else None
        self.chat_save_button.configure(state="normal" if self.last_assistant_answer else "disabled")
        self._refresh_chat_list(chat_id); self._update_attachment_label(); self._update_chat_mode_ui(); self._render_chat()

    def _mode_changed(self, _event=None):
        if self.chat_busy:
            active = next((item for item in self.chats if item.get("id") == self.active_chat_id), None)
            self.chat_mode_var.set(active.get("mode", "Обычный чат") if active else "Обычный чат")
            return
        self._update_chat_mode_ui(); self._save_current_chat()
        if self.chat_mode_var.get() == "Работа по цели": self.goal_entry.focus_set()

    def _update_chat_mode_ui(self):
        if self.chat_mode_var.get() == "Работа по цели": self.goal_bar.grid(row=2, column=0, sticky="ew")
        else: self.goal_bar.grid_remove()
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
        sidebar = ttk.Frame(outer, padding=8); sidebar.columnconfigure(0, weight=1); sidebar.rowconfigure(2, weight=1)
        main = ttk.Frame(outer); main.columnconfigure(0, weight=1); main.rowconfigure(3, weight=1)
        outer.add(sidebar, weight=1); outer.add(main, weight=4)
        ttk.Label(sidebar, text="История чатов", font=("Segoe UI", 11, "bold")).grid(row=0, column=0, sticky="w", pady=(0,6))
        ttk.Button(sidebar, text="＋ Новый чат", command=self.new_chat, style="Accent.TButton").grid(row=1, column=0, sticky="ew", pady=(0,6))
        self.chat_listbox = tk.Listbox(sidebar, exportselection=False, activestyle="dotbox", font=("Segoe UI", 10))
        self.chat_listbox.configure(background="#e9edf5", foreground="#334155", selectbackground="#c7d2fe", selectforeground="#1e1b4b", borderwidth=0, highlightthickness=0)
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
        self.chat_provider_box = ttk.Combobox(head, textvariable=self.provider, values=["Ollama", "LM Studio"], state="readonly", width=12)
        self.chat_provider_box.grid(row=0, column=4, padx=4); self.chat_provider_box.bind("<<ComboboxSelected>>", self._provider_changed)
        ttk.Button(head, text="Модели", command=self.check_models).grid(row=0, column=5, padx=4)
        folder_bar = ttk.Frame(main, padding=(8, 0, 8, 6)); folder_bar.grid(row=1, column=0, sticky="ew"); folder_bar.columnconfigure(1, weight=1)
        ttk.Button(folder_bar, text="Рабочая папка…", command=self.choose_chat_working_dir).grid(row=0, column=0, padx=(0,6))
        self.chat_folder_label = tk.StringVar(value=self.chat_working_dir or "Рабочая папка не выбрана")
        ttk.Label(folder_bar, textvariable=self.chat_folder_label).grid(row=0, column=1, sticky="ew")
        ttk.Button(folder_bar, text="Открыть папку", command=self.open_chat_working_dir).grid(row=0, column=2, padx=6)
        self.goal_bar = ttk.Frame(main, padding=(8, 0, 8, 6)); self.goal_bar.columnconfigure(1, weight=1)
        ttk.Label(self.goal_bar, text="Цель:").grid(row=0, column=0, sticky="w")
        self.goal_entry = ttk.Entry(self.goal_bar, textvariable=self.goal_var)
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
        ttk.Button(tree_actions, text="Обновить", command=self.refresh_project_files).pack(side="left", padx=2)
        ttk.Button(tree_actions, text="Прикрепить выбранный", command=self.attach_project_file).pack(side="left", padx=2)
        self.editor_panel = ttk.Frame(body, padding=6); self.editor_panel.columnconfigure(0, weight=1); self.editor_panel.rowconfigure(1, weight=1)
        self.project_file_label = tk.StringVar(value="Выберите файл проекта")
        ttk.Label(self.editor_panel, textvariable=self.project_file_label, style="Title.TLabel").grid(row=0, column=0, sticky="ew", pady=(0, 5))
        self.project_editor = tk.Text(self.editor_panel, wrap="none", state="disabled", font=("Cascadia Code", 10), background="#ffffff", foreground="#1f2937", insertbackground="#4f46e5", borderwidth=0, highlightthickness=1, highlightbackground="#d7ddea")
        self.project_editor.grid(row=1, column=0, sticky="nsew")
        editor_scroll = ttk.Scrollbar(self.editor_panel, orient="vertical", command=self.project_editor.yview); editor_scroll.grid(row=1, column=1, sticky="ns"); self.project_editor.configure(yscrollcommand=editor_scroll.set)
        editor_actions = ttk.Frame(self.editor_panel); editor_actions.grid(row=2, column=0, sticky="ew", pady=4)
        ttk.Button(editor_actions, text="Открыть в приложении", command=self.open_project_file).pack(side="left", padx=2)
        ttk.Button(editor_actions, text="Задача с файлом → diff", command=self.prepare_edit, style="Accent.TButton").pack(side="right", padx=2)
        transcript = ttk.Frame(body); transcript.columnconfigure(0, weight=1); transcript.rowconfigure(0, weight=1)
        body.add(transcript, weight=4); self.chat_body = body
        self.chat_display = tk.Text(transcript, wrap="word", state="disabled", font=("Segoe UI", 11), padx=20, pady=16, background="#ffffff", foreground="#1f2937", insertbackground="#4f46e5", selectbackground="#c7d2fe", borderwidth=0, highlightthickness=0)
        self.chat_display.grid(row=0, column=0, sticky="nsew")
        chat_history_scroll = ttk.Scrollbar(transcript, orient="vertical", command=self.chat_display.yview)
        chat_history_scroll.grid(row=0, column=1, sticky="ns"); self.chat_display.configure(yscrollcommand=chat_history_scroll.set)
        self.chat_display.tag_configure("user", foreground="#4338ca", spacing1=14, spacing3=7, font=("Segoe UI Semibold", 10, "bold"))
        self.chat_display.tag_configure("assistant", foreground="#1f2937", spacing1=9, spacing3=18)
        self.chat_display.tag_configure("system", foreground="#64748b", spacing1=8, spacing3=8)
        foot = ttk.Frame(main, padding=8); foot.grid(row=4, column=0, sticky="ew"); foot.columnconfigure(0, weight=1)
        self.attached_label = tk.StringVar(value="Файл не прикреплён")
        ttk.Label(foot, textvariable=self.attached_label).grid(row=0, column=0, sticky="w")
        attachment_actions = ttk.Frame(foot); attachment_actions.grid(row=0, column=1, padx=4)
        ttk.Button(attachment_actions, text="Прикрепить файлы…", command=self.attach_chat_file).pack(side="left", padx=2)
        ttk.Button(attachment_actions, text="Очистить вложения", command=self.clear_chat_files).pack(side="left", padx=2)
        self.chat_input = tk.Text(foot, height=4, wrap="word", font=("Segoe UI", 10), padx=10, pady=8, background="#ffffff", foreground="#1f2937", insertbackground="#4f46e5", selectbackground="#c7d2fe", borderwidth=1, relief="solid", highlightthickness=1, highlightbackground="#d7ddea", highlightcolor="#818cf8")
        self.chat_input.grid(row=1, column=0, sticky="ew", pady=6); self.chat_input.bind("<Return>", self._chat_enter)
        controls = ttk.Frame(foot); controls.grid(row=1, column=1, sticky="se", padx=4)
        self.chat_send_button = ttk.Button(controls, text="Отправить", command=self.send_chat, style="Accent.TButton"); self.chat_send_button.pack(fill="x", pady=2)
        self.chat_edit_button = ttk.Button(controls, text="Выполнить задачу с файлом", command=self.prepare_edit); self.chat_edit_button.pack(fill="x", pady=2)
        self.chat_save_button = ttk.Button(controls, text="Сохранить ответ как файл…", command=self.save_chat_answer, state="disabled"); self.chat_save_button.pack(fill="x", pady=2)
        ttk.Label(foot, text="Enter — отправить · Shift+Enter — новая строка · Файлы передаются локальной модели, сохранение ответа выполняется вами.").grid(row=2, column=0, columnspan=2, sticky="w")
        self._refresh_chat_list()
        if self.chats:
            latest = max(self.chats, key=lambda item: item.get("updated", 0))
            self._activate_chat(latest["id"])
        else:
            self.new_chat()
        self.chat_input.focus_set()

    def _chat_enter(self, event):
        if event.state & 0x0001: return
        self.send_chat(); return "break"

    def _render_chat(self):
        if not self.chat_window or not self.chat_window.winfo_exists(): return
        self.chat_display.configure(state="normal"); self.chat_display.delete("1.0", "end")
        if not self.chat_history:
            mode_hint = {"Работа по цели": "Опишите цель в поле выше — чат сохранит её и будет помогать двигаться к результату по шагам.",
                         "Режим Cursor": "Выберите рабочую папку проекта. Откройте файл в дереве, смотрите код и задавайте задачу. Изменения появятся в diff и применятся только после подтверждения.",
                         "Обычный чат": "Задавайте вопросы и прикрепляйте файлы из рабочей папки."}.get(self.chat_mode_var.get(), "Задавайте вопросы.")
            self.chat_display.insert("end", f"Локальный чат\n\n{mode_hint}\n\nМодель не имеет доступа к файловой системе и не может удалять файлы.", "system")
        for item in self.chat_history:
            label = "Вы" if item["role"] == "user" else "Локальная модель"
            self.chat_display.insert("end", label + "\n", item["role"])
            self.chat_display.insert("end", item["content"] + "\n\n", item["role"])
        self.chat_display.configure(state="disabled"); self.chat_display.see("end")

    def new_chat(self):
        if self.chat_busy: return
        self._save_current_chat()
        chat = {"id": uuid.uuid4().hex, "title": "Новый чат", "history": [], "attachments": [],
                "mode": "Обычный чат", "goal": "", "working_dir": self.chat_working_dir, "updated": time.time()}
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
            self._save_current_chat()
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
            if not self.chat_working_dir: return False, "сначала задайте рабочую папку проекта"
            try: path.relative_to(Path(self.chat_working_dir).resolve())
            except ValueError: return False, "в режиме кода разрешены только файлы из рабочей папки"
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
        self.attached_label.set("Вложения: " + (", ".join(names) if names else "нет"))

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
            if self.chat_working_dir: system += "\nРабочий проект: " + Path(self.chat_working_dir).name
            active_path = getattr(self, "project_file_path", None)
            if active_path and active_path.is_file() and not any(item.get("path") == str(active_path) for item in self.chat_attachments):
                try:
                    active_text = ("Selected image: " + active_path.name if active_path.suffix.lower() in IMAGE_EXTS else
                                   excel_workbook_text(active_path, 50000) if active_path.suffix.lower() in {".xlsx", ".xlsm", ".xls", ".xlsb"}
                                   else extract_text(active_path) if active_path.suffix.lower() in SUPPORTED
                                   else read_editable_text(active_path)[0][:50000] if is_plain_text_file(active_path)
                                   else "Binary file selected; no local content adapter is available.")
                    system += f"\n\n--- Открытый файл проекта: {active_path.name} ---\n{active_text}"
                except (OSError, UnicodeError, ValueError, ImportError, zipfile.BadZipFile): pass
        if self.chat_attachments:
            system += "\n\nПользователь передал содержимое файлов только для этой беседы. Рабочая папка чата: " + (Path(self.chat_working_dir).name if self.chat_working_dir else "не выбрана")
            for item in self.chat_attachments:
                system += f"\n\n--- Вложение: {item['name']} ---\n{item['text']}"
        messages = [{"role": "system", "content": system}] + self.chat_history
        image_paths = [item["path"] for item in self.chat_attachments if item.get("image") and Path(item.get("path", "")).is_file()]
        active_path = getattr(self, "project_file_path", None)
        if self.chat_mode_var.get() == "Режим Cursor" and active_path and active_path.suffix.lower() in IMAGE_EXTS and active_path.is_file():
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
        edit_words = re.search(r"\b(измени|изменить|исправь|исправить|добавь|добавить|заполни|заполнить|замени|заменить|перепиши|переписать|отредактируй|отредактировать|переведи|перевести|сократи|сократить|обнови|обновить|change|edit|update|rewrite|fill|translate|format)\b", prompt, re.IGNORECASE)
        active_path = (getattr(self, "project_file_path", None) if self.chat_mode_var.get() == "Режим Cursor" else None)
        candidate = active_path or self.chat_edit_path
        if not candidate:
            row = self.selected(); candidate = Path(row[0]) if row else None
        if edit_words and candidate and candidate.is_file():
            ext = candidate.suffix.lower()
            if ext in EDITABLE_TEXT | {".docx", ".xlsx", ".xlsm", ".xls", ".xlsb"} or (ext not in SUPPORTED and is_plain_text_file(candidate)):
                self.prepare_edit(); return
        self.chat_input.delete("1.0", "end"); self.chat_history.append({"role": "user", "content": prompt}); self._render_chat()
        self._save_current_chat()
        search_terms = extract_file_search_query(prompt)
        if search_terms is not None:
            results = self.index.find_relevant(search_terms)
            if results:
                lines = [f"Нашёл файлы в проиндексированных папках ({len(results)}):"]
                for path, name, ext, size, _mtime, _body in results:
                    lines.append(f"• {name} ({ext or 'без расширения'}, {human_size(size)})\n  {path}")
                answer = "\n".join(lines)
            else:
                answer = ("По этому запросу совпадений в локальном индексе нет. Поиск охватывает папки, добавленные во вкладке «Файлы и поиск» и просканированные кнопкой «Сканировать». "
                          "Добавьте нужный диск или папку и запустите сканирование, затем повторите запрос.")
            self.last_assistant_answer = answer
            if self.chat_window and self.chat_window.winfo_exists(): self.chat_save_button.configure(state="normal")
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
        if not row and not (self.chat_edit_path and self.chat_edit_path.is_file()) and not (self.chat_mode_var.get() == "Режим Cursor" and active_project_path and active_project_path.is_file()):
            messagebox.showinfo(APP_NAME, "Выберите файл в списке, затем попросите изменить его."); return
        if self.chat_mode_var.get() == "Режим Cursor" and active_project_path and active_project_path.is_file(): path = active_project_path
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
        self.chat_busy = True; self.chat_send_button.configure(state="disabled"); self.chat_edit_button.configure(state="disabled")
        provider = self.provider.get(); endpoint = self.endpoint.get().strip(); model = self.model.get().strip()
        self.config_data.update({"provider": provider, "endpoint": endpoint, "model": model}); write_config(self.config_data)
        self.status.set(f"Запрос к локальной модели {model}…")
        def work():
            try:
                answer = local_chat(provider, endpoint, model, messages)
                if not answer: raise ValueError("Модель вернула пустой ответ")
                self.after(0, lambda a=answer: self._chat_done(a, purpose, edit_path, original, edit_kind, edit_encoding))
            except Exception as e:
                self.after(0, lambda error=str(e): self._chat_error(error))
        threading.Thread(target=work, daemon=True).start()

    def _chat_done(self, answer, purpose, edit_path, original, edit_kind="text", edit_encoding="utf-8"):
        self.chat_busy = False
        if self.chat_window and self.chat_window.winfo_exists():
            self.chat_send_button.configure(state="normal"); self.chat_edit_button.configure(state="normal")
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
            if self.chat_window and self.chat_window.winfo_exists(): self.chat_save_button.configure(state="normal")
            self.chat_history.append({"role": "assistant", "content": answer}); self._render_chat()
            self._save_current_chat()

    def _chat_error(self, error):
        self.chat_busy = False
        if self.chat_window and self.chat_window.winfo_exists():
            self.chat_send_button.configure(state="normal"); self.chat_edit_button.configure(state="normal")
        self.status.set("Не удалось получить ответ")
        messagebox.showerror("Локальная модель недоступна", f"{error}\n\nПроверьте подключение и наличие модели в Ollama или LM Studio.")

    def _show_edit_preview(self, path, original, edited, edit_kind, edit_encoding="utf-8"):
        win = tk.Toplevel(self); win.title(f"Проверка правок — {path.name}"); win.geometry("1000x680"); win.columnconfigure(0, weight=1); win.rowconfigure(1, weight=1)
        ttk.Label(win, text="Проверьте каждую правку. Исходник пока не менялся; резервная копия создаётся только при замене.", padding=8).grid(row=0, column=0, sticky="w")
        diff = tk.Text(win, wrap="none", font=("Consolas", 10)); diff.grid(row=1, column=0, sticky="nsew", padx=8, pady=4)
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
        ttk.Button(buttons, text="Сохранить как копию", command=lambda: self._save_edit_copy(path, edited, win, edit_kind, edit_encoding)).pack(side="left", padx=4)
        if edit_kind != "excel" or path.suffix.lower() in {".xlsx", ".xlsm"}:
            ttk.Button(buttons, text="Заменить исходник с backup", command=lambda: self._replace_with_backup(path, edited, win, edit_kind, edit_encoding)).pack(side="left", padx=4)
        ttk.Button(buttons, text="Отмена", command=win.destroy).pack(side="left", padx=4)

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
            else: temporary.write_bytes(edited.encode(edit_encoding))
            temporary.replace(path)
            self.index.refresh_path(path)
            messagebox.showinfo(APP_NAME, f"Изменения сохранены. Резервная копия:\n{backup}")
            window.destroy(); self._refresh_results()
        except (OSError, UnicodeError, ValueError, zipfile.BadZipFile, ET.ParseError) as e:
            messagebox.showerror(APP_NAME, f"Не удалось сохранить правки. Резервная копия (если создана):\n{backup}\n\n{e}")
    def _answer(self, answer):
        win = tk.Toplevel(self); win.title("Ответ локального ИИ"); win.geometry("720x520")
        text = tk.Text(win, wrap="word", padx=12, pady=12); text.pack(fill="both", expand=True); text.insert("1.0", answer); text.configure(state="disabled")
    def _poll(self):
        try:
            while True: self.status.set(self.events.get_nowait()); self._refresh_results()
        except queue.Empty: pass
        self.after(500, self._poll)


if __name__ == "__main__":
    # Needed for the standard-library simpledialog helper on all Tk builds.
    from tkinter import simpledialog
    App().mainloop()
