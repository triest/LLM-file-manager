"""Local-first file manager prototype. Python 3.10+, standard library only."""
from __future__ import annotations

import json
import os
import queue
import re
import shutil
import sqlite3
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.request
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


def app_data() -> Path:
    root = Path(os.environ.get("APPDATA", Path.home())) / "LocalFileManager"
    root.mkdir(parents=True, exist_ok=True)
    return root


def read_config() -> dict:
    p = app_data() / "config.json"
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {"folders": [], "model": "qwen2.5:7b"}


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
                clauses = " AND ".join(["(name LIKE ? OR body LIKE ? OR path LIKE ?)"] * len(words))
                args = [v for word in words for v in (f"%{word}%", f"%{word}%", f"%{word}%")]
                rows = self.db.execute(f"SELECT path,name,ext,size,mtime,body FROM files WHERE {clauses} ORDER BY mtime DESC LIMIT ?", (*args, limit)).fetchall() if words else []
        return rows

    def counts(self):
        with self.lock:
            return self.db.execute("SELECT COUNT(*), COALESCE(SUM(size),0) FROM files").fetchone()


def ollama_request(payload: dict, timeout=120):
    req = urllib.request.Request(OLLAMA_URL + "/api/generate", data=json.dumps(payload).encode(), headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as r: return json.loads(r.read())


def ollama_tags():
    with urllib.request.urlopen(OLLAMA_URL + "/api/tags", timeout=2) as r: return [m["name"] for m in json.loads(r.read()).get("models", [])]


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
        model_bar = ttk.Frame(right); model_bar.grid(row=3, column=0, sticky="ew"); model_bar.columnconfigure(0, weight=1)
        self.model = tk.StringVar(value=self.config_data.get("model", "qwen2.5:7b"))
        self.model_box = ttk.Combobox(model_bar, textvariable=self.model, values=[self.model.get(), "qwen2.5:3b", "qwen2.5:7b", "deepseek-r1:14b", "qwen2.5vl:7b", "qwen3-vl:4b"]); self.model_box.grid(row=0, column=0, sticky="ew")
        ttk.Button(model_bar, text="Подключить", command=self.check_models).grid(row=0, column=1, padx=4)
        self.ai_question = tk.StringVar(); ttk.Entry(right, textvariable=self.ai_question).grid(row=4, column=0, sticky="ew", pady=5); self.ai_question.set("Спросить о найденных документах…")
        ttk.Button(right, text="Ответить по найденным файлам", command=self.ask_ai).grid(row=5, column=0, sticky="ew")
        ttk.Button(right, text="Предложить порядок для файлов", command=self.organize_ai).grid(row=6, column=0, sticky="ew", pady=5)
        ttk.Label(right, text="* PDF: для извлечения текста установите pypdf (необязательно).\nИИ работает локально через Ollama, если он установлен.", wraplength=370).grid(row=7, column=0, sticky="w", pady=5)
        self.status = tk.StringVar(value="Добавьте папку и нажмите «Сканировать»."); ttk.Label(self, textvariable=self.status, relief="sunken", anchor="w", padding=4).grid(row=2, column=0, sticky="ew")

    def _refresh_folders(self): self.folder_var.set("; ".join(self.config_data["folders"]) or "Папки не выбраны")
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
        try: Path(row[0]).rename(target); self._refresh_results()
        except OSError as e: messagebox.showerror(APP_NAME, str(e))
    def move_selected(self):
        row = self.selected()
        if not row: return
        folder = filedialog.askdirectory(title="Куда переместить файл")
        if not folder: return
        source = Path(row[0]); target = Path(folder) / source.name
        if target.exists(): messagebox.showerror(APP_NAME, "Файл с таким именем уже есть в папке."); return
        try: shutil.move(str(source), str(target)); self._refresh_results()
        except OSError as e: messagebox.showerror(APP_NAME, str(e))
    def check_models(self):
        self.status.set("Проверка локального Ollama…")
        def work():
            try:
                tags = ollama_tags()
                if not tags and try_start_ollama():
                    for _ in range(15):
                        time.sleep(1)
                        try: tags = ollama_tags(); break
                        except Exception: pass
                self.events.put("Модели Ollama: " + (", ".join(tags) if tags else "сервер не найден. Установите Ollama и скачайте модель."))
                if tags: self.after(0, lambda: self.model_box.configure(values=tags))
            except Exception:
                started = try_start_ollama()
                self.events.put("Запускаю Ollama…" if started else "Ollama не найдена. Установите Ollama; скачанные модели уже обнаружены в локальном хранилище.")
        threading.Thread(target=work, daemon=True).start()
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
        model = self.model.get().strip(); self.config_data["model"] = model; write_config(self.config_data)
        self.status.set("ИИ обрабатывает запрос локально…")
        def work():
            try:
                result = ollama_request({"model": model, "prompt": prompt, "stream": False, "options": {"temperature": 0.2}})
                answer = result.get("response", "Пустой ответ")
                self.after(0, lambda: self._answer(answer))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Локальная модель недоступна", f"{e}\n\nУбедитесь, что Ollama установлена, запущена и выбранная модель скачана."))
            self.events.put("Готово")
        threading.Thread(target=work, daemon=True).start()
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
