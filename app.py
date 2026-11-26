"""Local-first file manager prototype. Python 3.10+, standard library only."""
from __future__ import annotations

import json
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
SUPPORTED = {".txt", ".md", ".rst", ".csv", ".tsv", ".json", ".xml", ".html", ".htm", ".py", ".js", ".ts", ".css", ".log", ".ini", ".yaml", ".yml", ".toml", ".rtf", ".docx", ".xlsx", ".pptx", ".odt", ".ods", ".odp", ".pdf"}
MAX_TEXT = 500_000
OLLAMA_URL = "http://127.0.0.1:11434"
EDITABLE_TEXT = {".txt", ".md", ".rst", ".csv", ".tsv", ".json", ".xml", ".html", ".htm", ".py", ".js", ".ts", ".css", ".log", ".ini", ".yaml", ".yml", ".toml"}
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


def extract_json_object(response: str):
    cleaned = re.sub(r"<think>.*?</think>", "", response, flags=re.DOTALL).strip()
    start = cleaned.find("{")
    if start < 0: raise ValueError("В ответе нет JSON")
    result, _end = json.JSONDecoder().raw_decode(cleaned[start:])
    return result


def extract_text(path: Path) -> str:
    ext = path.suffix.lower()
    try:
        if ext in {".docx", ".xlsx", ".pptx", ".odt", ".ods", ".odp"}:
            chunks = []
            with zipfile.ZipFile(path) as archive:
                names = archive.namelist()
                if ext == ".docx": names = [n for n in names if re.fullmatch(r"word/(document|header\d+|footer\d+|footnotes|endnotes)\.xml", n)]
                elif ext == ".xlsx": names = [n for n in names if n == "xl/sharedStrings.xml" or re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)]
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
    except (OSError, ValueError, zipfile.BadZipFile, KeyError, RuntimeError):
        return ""


class Index:
    def __init__(self):
        self.db = sqlite3.connect(app_data() / "files.sqlite3", check_same_thread=False)
        self.db.create_function("CASEFOLD", 1, lambda value: str(value or "").casefold())
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
                args = [v for word in words for v in (f"%{word.casefold()}%", f"%{word.casefold()}%", f"%{word.casefold()}%")]
                rows = self.db.execute(f"SELECT path,name,ext,size,mtime,body FROM files WHERE {clauses} ORDER BY mtime DESC LIMIT ?", (*args, limit)).fetchall() if words else []
        return rows

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
    payload = {"model": model, "messages": messages, "stream": False}
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


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME); self.geometry("1180x760"); self.minsize(900, 600)
        self.config_data = read_config(); self.index = Index(); self.events = queue.Queue(); self.rows = []
        self.provider = tk.StringVar(value=self.config_data.get("provider", "Ollama"))
        self.endpoint = tk.StringVar(value=self.config_data.get("endpoint", OLLAMA_URL))
        self.chat_history = []; self.chat_context = ""; self.chat_window = None; self.chat_busy = False
        self._build(); self._refresh_folders(); self._refresh_results(); self._poll()
        # Start/check the local Ollama service automatically; no console is shown.
        self.after(900, self.check_models)

    def _build(self):
        self.columnconfigure(0, weight=1); self.rowconfigure(1, weight=1)
        bar = ttk.Frame(self, padding=10); bar.grid(row=0, column=0, sticky="ew"); bar.columnconfigure(1, weight=1)
        ttk.Label(bar, text="Папки для индексации:").grid(row=0, column=0, sticky="w")
        self.folder_var = tk.StringVar(); ttk.Entry(bar, textvariable=self.folder_var, state="readonly").grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(bar, text="Добавить папку", command=self.add_folder).grid(row=0, column=2, padx=3)
        ttk.Button(bar, text="Убрать", command=self.remove_folder).grid(row=0, column=3, padx=3)
        ttk.Button(bar, text="Сканировать", command=self.scan).grid(row=0, column=4, padx=3)
        body = ttk.Panedwindow(self, orient="horizontal"); body.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0,10))
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
        except (OSError, UnicodeError, zipfile.BadZipFile, ET.ParseError) as e: messagebox.showerror(APP_NAME, str(e))
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
        if self.chat_window and self.chat_window.winfo_exists():
            self.chat_window.lift(); return
        win = self.chat_window = tk.Toplevel(self)
        win.title("Локальный ИИ — чат"); win.geometry("850x720"); win.minsize(650, 520)
        win.columnconfigure(0, weight=1); win.rowconfigure(1, weight=1)
        head = ttk.Frame(win, padding=8); head.grid(row=0, column=0, sticky="ew"); head.columnconfigure(1, weight=1)
        ttk.Label(head, text="Модель:").grid(row=0, column=0)
        self.chat_model_box = ttk.Combobox(head, textvariable=self.model, values=self.model_box.cget("values"), width=22)
        self.chat_model_box.grid(row=0, column=1, sticky="ew", padx=6); head.columnconfigure(1, weight=1)
        self.chat_provider_box = ttk.Combobox(head, textvariable=self.provider, values=["Ollama", "LM Studio"], state="readonly", width=12)
        self.chat_provider_box.grid(row=0, column=2, padx=4); self.chat_provider_box.bind("<<ComboboxSelected>>", self._provider_changed)
        ttk.Button(head, text="Найти модели", command=self.check_models).grid(row=0, column=3, padx=4)
        ttk.Button(head, text="Новый чат", command=self.new_chat).grid(row=0, column=4, padx=4)
        self.chat_display = tk.Text(win, wrap="word", state="disabled", font=("Segoe UI", 11), padx=14, pady=12)
        self.chat_display.grid(row=1, column=0, sticky="nsew", padx=8)
        self.chat_display.tag_configure("user", foreground="#174ea6", spacing1=12, spacing3=10, font=("Segoe UI", 10, "bold"))
        self.chat_display.tag_configure("assistant", foreground="#202124", spacing1=10, spacing3=18)
        self.chat_display.tag_configure("system", foreground="#5f6368", spacing1=8, spacing3=8)
        foot = ttk.Frame(win, padding=8); foot.grid(row=2, column=0, sticky="ew"); foot.columnconfigure(0, weight=1)
        self.attached_label = tk.StringVar(value="Файл не прикреплён")
        ttk.Label(foot, textvariable=self.attached_label).grid(row=0, column=0, sticky="w")
        ttk.Button(foot, text="Прикрепить выбранный файл", command=self.attach_chat_file).grid(row=0, column=1, padx=4)
        self.chat_input = tk.Text(foot, height=4, wrap="word", font=("Segoe UI", 10))
        self.chat_input.grid(row=1, column=0, sticky="ew", pady=6); self.chat_input.bind("<Return>", self._chat_enter)
        controls = ttk.Frame(foot); controls.grid(row=1, column=1, sticky="se", padx=4)
        self.chat_send_button = ttk.Button(controls, text="Отправить", command=self.send_chat); self.chat_send_button.pack(fill="x", pady=2)
        self.chat_edit_button = ttk.Button(controls, text="Подготовить правку", command=self.prepare_edit); self.chat_edit_button.pack(fill="x", pady=2)
        ttk.Label(foot, text="Enter — отправить · Shift+Enter — новая строка · Правки применяются только после проверки и подтверждения.").grid(row=2, column=0, columnspan=2, sticky="w")
        self.chat_history = []; self.chat_context = ""; self._render_chat()
        self.chat_input.focus_set()

    def _chat_enter(self, event):
        if event.state & 0x0001: return
        self.send_chat(); return "break"

    def _render_chat(self):
        if not self.chat_window or not self.chat_window.winfo_exists(): return
        self.chat_display.configure(state="normal"); self.chat_display.delete("1.0", "end")
        if not self.chat_history:
            self.chat_display.insert("end", "Локальный чат\n\nВыберите локальную модель выше. Можно задавать вопросы, прикреплять файлы и просить модель подготовить правки.\n\nМодель не имеет доступа к файловой системе и не может удалять файлы.", "system")
        for item in self.chat_history:
            label = "Вы" if item["role"] == "user" else "Локальная модель"
            self.chat_display.insert("end", label + "\n", item["role"])
            self.chat_display.insert("end", item["content"] + "\n\n", item["role"])
        self.chat_display.configure(state="disabled"); self.chat_display.see("end")

    def new_chat(self):
        self.chat_history = []; self.chat_context = ""; self.attached_label.set("Файл не прикреплён"); self._render_chat()

    def attach_chat_file(self):
        row = self.selected()
        if not row:
            messagebox.showinfo(APP_NAME, "Сначала выберите файл в списке."); return
        if not row[5]:
            messagebox.showinfo(APP_NAME, "Для этого типа файла не извлечён текст."); return
        self.chat_context = f"Прикреплённый файл: {row[0]}\nТекст файла:\n{row[5][:30000]}"
        self.attached_label.set(Path(row[0]).name + " — текст передаётся только локальной модели")

    def _chat_messages(self):
        system = "Ты помощник в локальном файловом менеджере. Отвечай на языке пользователя. У тебя нет инструментов и доступа к диску: ты не можешь перемещать, создавать или удалять файлы. Для правки верни полный текст изменённого файла по явному запросу пользователя; приложение покажет его и попросит подтверждение."
        if self.chat_context: system += "\n\n" + self.chat_context
        return [{"role": "system", "content": system}] + self.chat_history

    def send_chat(self):
        if self.chat_busy: return
        prompt = self.chat_input.get("1.0", "end-1c").strip()
        if not prompt: return
        self.chat_input.delete("1.0", "end"); self.chat_history.append({"role": "user", "content": prompt}); self._render_chat()
        self._run_chat(self._chat_messages(), purpose="chat")

    def prepare_edit(self):
        if self.chat_busy: return
        row = self.selected()
        if not row:
            messagebox.showinfo(APP_NAME, "Выберите файл в списке, затем попросите изменить его."); return
        path = Path(row[0]); ext = path.suffix.lower()
        edit_kind = "docx" if ext == ".docx" else "text"
        if ext not in EDITABLE_TEXT and edit_kind != "docx":
            messagebox.showinfo(APP_NAME, "Правки поддерживаются для текста, исходного кода и Word DOCX. Для других Office-файлов доступен чат по содержимому, без перезаписи книги или презентации."); return
        try:
            if edit_kind == "docx":
                original = docx_paragraphs(path)
                if not original: messagebox.showinfo(APP_NAME, "В документе не найдено текстовых абзацев."); return
                if len(original) > 120 or sum(len(p["text"]) for p in original) > 30000:
                    messagebox.showinfo(APP_NAME, "Для правки в прототипе выберите DOCX до 120 абзацев и 30 000 символов."); return
            else:
                original, edit_encoding = read_editable_text(path)
                if len(original) > 50000: messagebox.showinfo(APP_NAME, "Для правки выберите текстовый файл не длиннее 50 000 символов."); return
        except (OSError, UnicodeError, zipfile.BadZipFile, ET.ParseError) as e:
            messagebox.showerror(APP_NAME, f"Не удалось прочитать файл: {e}"); return
        request = self.chat_input.get("1.0", "end-1c").strip()
        if not request:
            messagebox.showinfo(APP_NAME, "Напишите, что нужно изменить, в поле сообщения."); return
        if edit_kind == "docx":
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
        self.chat_history.append({"role": "user", "content": f"Подготовить правку `{path.name}`: {request}"}); self._render_chat()
        self._run_chat([{"role": "system", "content": "Ты редактируешь только переданный текст. Не управляешь файловой системой. Выполняй только просьбу пользователя."}, {"role": "user", "content": instruction}], purpose="edit", edit_path=path, original=original, edit_kind=edit_kind, edit_encoding=locals().get("edit_encoding", "utf-8"))

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
            if edit_kind == "docx":
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
            self.chat_history.append({"role": "assistant", "content": "Подготовлены правки для проверки. Исходный файл пока не изменён."}); self._render_chat()
            self._show_edit_preview(edit_path, original, edited, edit_kind, edit_encoding)
        else:
            self.chat_history.append({"role": "assistant", "content": answer}); self._render_chat()

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
        if edit_kind == "docx":
            preview = []
            for change in edited:
                preview.extend(difflib.unified_diff([change["text"] + "\n"], [change["new_text"] + "\n"], fromfile=f"абзац {change['id']} — исходник", tofile=f"абзац {change['id']} — правка"))
            diff.insert("1.0", "".join(preview) or "Изменений нет.")
        else:
            diff.insert("1.0", "".join(difflib.unified_diff(original.splitlines(True), edited.splitlines(True), fromfile="исходник", tofile="версия ИИ")) or "Изменений нет.")
        diff.configure(state="disabled")
        buttons = ttk.Frame(win, padding=8); buttons.grid(row=2, column=0, sticky="e")
        ttk.Button(buttons, text="Сохранить как копию", command=lambda: self._save_edit_copy(path, edited, win, edit_kind, edit_encoding)).pack(side="left", padx=4)
        ttk.Button(buttons, text="Заменить исходник с backup", command=lambda: self._replace_with_backup(path, edited, win, edit_kind, edit_encoding)).pack(side="left", padx=4)
        ttk.Button(buttons, text="Отмена", command=win.destroy).pack(side="left", padx=4)

    def _save_edit_copy(self, path, edited, window, edit_kind, edit_encoding="utf-8"):
        target = path.with_name(path.stem + "_edited" + path.suffix); number = 2
        while target.exists():
            target = path.with_name(f"{path.stem}_edited_{number}{path.suffix}"); number += 1
        try:
            if edit_kind == "docx": write_docx_edits(path, target, edited)
            else: target.write_bytes(edited.encode(edit_encoding))
            self.index.refresh_path(target); messagebox.showinfo(APP_NAME, f"Копия сохранена:\n{target}"); window.destroy(); self._refresh_results()
        except (OSError, UnicodeError, zipfile.BadZipFile, ET.ParseError) as e: messagebox.showerror(APP_NAME, str(e))

    def _replace_with_backup(self, path, edited, window, edit_kind, edit_encoding="utf-8"):
        if not messagebox.askyesno("Подтвердить правку", f"Применить изменения к {path.name}? Перед правкой будет создана резервная копия исходного файла.\n\nМодель не может удалять файлы."):
            return
        backup = path.with_name(path.name + ".backup-" + datetime.now().strftime("%Y%m%d-%H%M%S-%f"))
        temporary = path.with_name(f".{path.name}.{uuid.uuid4().hex}.tmp")
        try:
            shutil.copy2(path, backup)
            if edit_kind == "docx": write_docx_edits(path, temporary, edited)
            else: temporary.write_bytes(edited.encode(edit_encoding))
            temporary.replace(path)
            self.index.refresh_path(path)
            messagebox.showinfo(APP_NAME, f"Изменения сохранены. Резервная копия:\n{backup}")
            window.destroy(); self._refresh_results()
        except (OSError, UnicodeError, zipfile.BadZipFile, ET.ParseError) as e:
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
