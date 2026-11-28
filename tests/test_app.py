import json
import logging
import tempfile
import threading
import time
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

import app


class TextAndSearchTests(unittest.TestCase):
    def test_keyboard_layout_search_keeps_subject_terms(self):
        self.assertEqual(app.extract_file_search_query("Yfqlb nfv afqk c rjynhfrnfvb"), ["контрактами"])

    def test_text_extractors_handle_html_and_bad_office_archives(self):
        with tempfile.TemporaryDirectory() as directory:
            html = Path(directory) / "page.html"
            html.write_text("<title>Index</title><p>Привет &amp; мир</p>", encoding="utf-8")
            self.assertIn("Привет & мир", app.extract_text(html))
            broken = Path(directory) / "broken.docx"
            broken.write_text("not a zip", encoding="utf-8")
            self.assertEqual(app.extract_text(broken), "")

    def test_index_scans_selected_roots_and_finds_russian_terms(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "selected"
            root.mkdir()
            (root / "контракт.txt").write_text("Договор поставки и сроки оплаты", encoding="utf-8")
            hidden = root / ".git"
            hidden.mkdir()
            (hidden / "secret.txt").write_text("Не индексировать", encoding="utf-8")
            data = Path(directory) / "data"
            data.mkdir()
            with patch.object(app, "app_data", return_value=data):
                index = app.Index()
                try:
                    progress = []
                    index.scan([str(root)], progress.append)
                    rows = index.find_relevant(["контрактами"])
                    self.assertEqual(len(rows), 1)
                    self.assertIn("контракт.txt", rows[0][1])
                    self.assertEqual(index.counts()[0], 1)
                    self.assertTrue(progress[-1].startswith("Индексация завершена"))
                finally:
                    index.db.close()

    def test_model_manifest_and_direct_gguf_discovery(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest_path = root / "manifests" / "registry.ollama.ai" / "library" / "qwen2.5" / "7b"
            manifest_path.parent.mkdir(parents=True)
            blob = root / "blobs" / ("sha256-" + "a" * 64)
            blob.parent.mkdir()
            blob.write_bytes(b"gguf")
            manifest_path.write_text(json.dumps({"layers": [{"mediaType": "application/vnd.ollama.image.model", "digest": "sha256:" + "a" * 64}]}), encoding="utf-8")
            direct = root / "coder.gguf"
            direct.write_bytes(b"gguf")
            found = app.discover_model_files(root)
            self.assertEqual(found["qwen2.5:7b"], blob.resolve())
            self.assertEqual(found["coder"], direct.resolve())


class ContextTests(unittest.TestCase):
    def test_compaction_drops_old_runtime_errors_and_keeps_latest_request(self):
        messages = [
            {"role": "system", "content": "rules"},
            {"role": "user", "content": "old question " * 40},
            {"role": "assistant", "content": "Ошибка локальной модели: secret"},
            {"role": "user", "content": "current attachment content"},
        ]
        compacted, original, sent = app.compact_chat_messages(messages, max_chars=120)
        self.assertEqual(compacted[-1]["content"], "current attachment content")
        self.assertNotIn("secret", str(compacted))
        self.assertEqual(original, len("rules") + len("old question " * 40) + len("current attachment content"))
        self.assertLessEqual(sent, 120)

    def test_compaction_explains_overflow_without_truncating_current_message(self):
        current = "x" * 30
        with self.assertRaisesRegex(ValueError, "безопасного лимита"):
            app.compact_chat_messages([{"role": "user", "content": current}], max_chars=20)

    def test_reasoning_markers_are_separated_from_final_answer(self):
        final, reasoning = app.split_model_reasoning("размышление [End thinking] готовый ответ")
        self.assertEqual(reasoning, "размышление")
        self.assertEqual(final, "готовый ответ")


class DocumentEditTests(unittest.TestCase):
    def test_excel_edit_is_validated_written_and_detects_stale_preview(self):
        from openpyxl import Workbook, load_workbook

        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.xlsx"
            target = Path(directory) / "edited.xlsx"
            workbook = Workbook()
            workbook.active.title = "Смета"
            workbook.active["A1"] = "Старое"
            workbook.save(source)
            workbook.close()

            plan = {"operations": [{"sheet": "Смета", "cell": "A1", "value": "Новое"}]}
            validated = app.validate_excel_operations(source, plan)
            app.write_excel_operations(source, target, validated)
            result = load_workbook(target, data_only=False)
            try:
                self.assertEqual(result["Смета"]["A1"].value, "Новое")
            finally:
                result.close()

            workbook = load_workbook(source)
            workbook["Смета"]["A1"] = "Внешнее изменение"
            workbook.save(source)
            workbook.close()
            with self.assertRaisesRegex(ValueError, "изменилась после предпросмотра"):
                app.write_excel_operations(source, target, validated)

    def test_excel_rejects_duplicate_and_merged_cells(self):
        from openpyxl import Workbook

        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.xlsx"
            workbook = Workbook()
            workbook.active.merge_cells("A1:B1")
            workbook.save(source)
            workbook.close()
            duplicate = {"operations": [
                {"sheet": "Sheet", "cell": "C1", "value": 1},
                {"sheet": "Sheet", "cell": "c1", "value": 2},
            ]}
            with self.assertRaisesRegex(ValueError, "Повторная правка"):
                app.validate_excel_operations(source, duplicate)
            merged = {"operations": [{"sheet": "Sheet", "cell": "B1", "value": 1}]}
            with self.assertRaisesRegex(ValueError, "объединённый диапазон"):
                app.validate_excel_operations(source, merged)

    def test_docx_rewrite_preserves_other_package_parts(self):
        document_xml = (f'<w:document xmlns:w="{app.W_NS}"><w:body><w:p><w:r><w:t>До</w:t></w:r>'
                        '<w:r><w:t> правки</w:t></w:r></w:p></w:body></w:document>').encode()
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.docx"
            target = Path(directory) / "target.docx"
            with zipfile.ZipFile(source, "w") as archive:
                archive.writestr("word/document.xml", document_xml)
                archive.writestr("custom/data.bin", b"preserve")
            paragraphs = app.docx_paragraphs(source)
            self.assertEqual(paragraphs[0]["text"], "До правки")
            app.write_docx_edits(source, target, [{**paragraphs[0], "new_text": "После"}])
            with zipfile.ZipFile(target) as archive:
                self.assertEqual(archive.read("custom/data.bin"), b"preserve")
            self.assertEqual(app.docx_paragraphs(target)[0]["text"], "После")


class RuntimeTests(unittest.TestCase):
    def _run_fake_runtime(self, returncode=0, stdout="Проверочный ответ\nExiting...", stderr=""):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        model_path = Path(temp.name) / "model.gguf"
        model_path.write_bytes(b"fake model")
        commands = []
        logs = []

        class FakeProcess:
            def __init__(self, command, **_kwargs):
                commands.append(command)
                self.returncode = returncode

            def poll(self):
                return self.returncode

            def communicate(self, **_kwargs):
                return stdout, stderr

            def kill(self):
                self.returncode = -9

            def terminate(self):
                self.returncode = -15

        run_result = __import__("subprocess").CompletedProcess([], 0, "CUDA0: NVIDIA GPU\n", "")
        app._INFERENCE_STOP.clear()
        self.addCleanup(app._INFERENCE_STOP.clear)
        with patch.object(app, "runtime_chat_executable", return_value=Path("llama-completion.exe")), \
             patch.object(app, "resolve_model_path", return_value=model_path), \
             patch.object(app, "app_data", return_value=Path(temp.name) / "appdata"), \
             patch.object(app, "runtime_environment", return_value={}), \
             patch.object(app.subprocess, "run", return_value=run_result), \
             patch.object(app.subprocess, "Popen", side_effect=FakeProcess), \
             patch.object(app, "log_event", side_effect=logs.append):
            try:
                answer = app.local_chat(temp.name, "test-model", [{"role": "user", "content": "safe test question"}], timeout=2)
            except RuntimeError:
                answer = None
        return answer, commands, logs

    def test_runtime_explicitly_uses_cuda_and_preserved_context_configuration(self):
        answer, commands, _logs = self._run_fake_runtime()
        self.assertEqual(answer, "Проверочный ответ")
        args = commands[0]
        self.assertEqual(args[args.index("-dev") + 1], "CUDA0")
        self.assertEqual(args[args.index("-c") + 1], "65536")
        self.assertEqual(args[args.index("--rope-scaling") + 1], "yarn")
        self.assertEqual(args[args.index("--rope-scale") + 1], "2")
        self.assertEqual(args[args.index("--yarn-orig-ctx") + 1], "32768")
        self.assertEqual(args[args.index("-ctk") + 1], "q4_0")
        self.assertEqual(args[args.index("-ctv") + 1], "q4_0")
        self.assertEqual(args[args.index("-fa") + 1], "on")

    def test_runtime_errors_do_not_write_model_or_prompt_text_to_logs(self):
        secret = "SENSITIVE-ATTACHMENT-TEXT"
        _answer, _commands, logs = self._run_fake_runtime(returncode=2, stdout=secret, stderr=secret)
        self.assertNotIn(secret, "\n".join(logs))
        self.assertTrue(any("LLM_REQUEST_ERROR" in entry for entry in logs))


class TkWorkflowTests(unittest.TestCase):
    def test_all_chat_tabs_send_and_attachments_remain_complete(self):
        import tkinter as tk

        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        data = Path(temp.name) / "appdata"
        data.mkdir()
        original_handlers = list(logging.getLogger("local_file_manager").handlers)
        app._APP_LOGGER = None

        class FakeLocalChat:
            calls = []
            slow_request_started = threading.Event()
            release_slow_request = threading.Event()

            def __call__(self, _models_dir, _model, messages, **_kwargs):
                self.calls.append(messages)
                if messages[-1]["content"] == "slow flow":
                    self.slow_request_started.set()
                    self.release_slow_request.wait(5)
                return "reply"

        fake_chat = FakeLocalChat()
        try:
            with patch.object(app, "app_data", return_value=data), \
                 patch.object(app, "discover_model_files", return_value={}), \
                 patch.object(app.App, "_ensure_runtime", lambda _self: None), \
                 patch.object(app, "local_chat", side_effect=fake_chat), \
                 patch.object(app.messagebox, "showerror"):
                window = app.App()
                try:
                    window.withdraw()
                    self.assertEqual(len(window.main_notebook.tabs()), 3)
                    window.new_chat()

                    for tab, message, mode in ((window.chat_tab, "chat flow", "Обычный чат"),
                                               (window.files_tab, "files flow", "Обычный чат"),
                                               (window.cursor_tab, "cursor flow", "Режим Cursor"),
                                               (window.chat_tab, "goal flow", "Работа по цели")):
                        window.main_notebook.select(tab)
                        window.update()
                        window.new_chat()
                        window.chat_mode_var.set(mode)
                        if mode == "Работа по цели":
                            window.goal_var.set("Подтвердить работу режима цели")
                        window.chat_input.delete("1.0", "end")
                        window.chat_input.insert("1.0", message)
                        window.send_chat()
                        deadline = time.monotonic() + 8

                        def finish_when_ready():
                            if not window.chat_busy or time.monotonic() >= deadline:
                                window.quit()
                            else:
                                window.after(10, finish_when_ready)

                        window.after(10, finish_when_ready)
                        window.mainloop()
                        self.assertFalse(window.chat_busy, f"chat did not finish: {message}")
                        self.assertEqual(window.chat_history[-1]["content"], "reply")

                    window.main_notebook.select(window.chat_tab)
                    window.update()
                    window.new_chat()
                    originating_id = window.active_chat_id
                    window.chat_input.insert("1.0", "slow flow")
                    window.send_chat()

                    def switch_during_generation():
                        if fake_chat.slow_request_started.is_set():
                            window.new_chat()
                            fake_chat.release_slow_request.set()
                            return
                        window.after(10, switch_during_generation)

                    window.after(10, switch_during_generation)
                    deadline = time.monotonic() + 8

                    def finish_slow_request():
                        if not window.chat_busy or time.monotonic() >= deadline:
                            window.quit()
                        else:
                            window.after(10, finish_slow_request)

                    window.after(10, finish_slow_request)
                    window.mainloop()
                    origin_chat = next(item for item in window.chats if item["id"] == originating_id)
                    self.assertTrue(origin_chat["unread"])
                    self.assertEqual(origin_chat["history"][-1]["content"], "reply")
                    window._activate_chat(originating_id)
                    self.assertFalse(origin_chat["unread"])

                    self.assertIn("Подтвердить работу режима цели", fake_chat.calls[3][0]["content"])

                    window.main_notebook.select(window.files_tab)
                    window.update()
                    searchable = Path(temp.name) / "контракт.txt"
                    searchable.write_text("договор поставки", encoding="utf-8")
                    window.index.refresh_path(searchable)
                    window.new_chat()
                    window.files_chat_input.insert("1.0", "Yfqlb nfv afqk c rjynhfrnfvb")
                    call_count = len(fake_chat.calls)
                    window.send_chat()
                    self.assertEqual(len(fake_chat.calls), call_count)
                    self.assertIn("контракт.txt", window.chat_history[-1]["content"])

                    window.main_notebook.select(window.chat_tab)
                    window.update()
                    window.new_chat()
                    self.assertEqual(len(fake_chat.calls), 5)
                    self.assertEqual(len({id(window.main_chat_input), id(window.files_chat_input), id(window.cursor_chat_input)}), 3)

                    small = Path(temp.name) / "small.txt"
                    small.write_text("complete attachment", encoding="utf-8")
                    ok, reason = window._attach_path(small)
                    self.assertTrue(ok, reason)
                    window.chat_history = [{"role": "user", "content": "question"}]
                    messages = window._chat_messages()
                    self.assertIn("complete attachment", messages[-1]["content"])
                    self.assertNotIn("complete attachment", messages[0]["content"])

                    large = Path(temp.name) / "too-large.txt"
                    large.write_text("x" * (app.MAX_REQUEST_CHARS + 1), encoding="utf-8")
                    ok, reason = window._attach_path(large)
                    self.assertFalse(ok)
                    self.assertIn("не обрезано и не прикреплено", reason)
                    self.assertFalse(any(item["path"] == str(large.resolve()) for item in window.chat_attachments))
                    window.chat_attachments = [{"name": "old.txt", "path": str(small), "text": "partial", "truncated": True}]
                    window.chat_input.insert("1.0", "do not send partial file")
                    calls_before_partial = len(fake_chat.calls)
                    with patch.object(app.messagebox, "showerror"):
                        window.send_chat()
                    self.assertEqual(len(fake_chat.calls), calls_before_partial)
                    self.assertEqual(window.chat_input.get("1.0", "end-1c"), "do not send partial file")
                    window.chat_attachments = []

                    source = Path(temp.name) / "reviewed.txt"
                    source.write_text("before", encoding="utf-8")
                    window._show_edit_preview(source, "before", "after", "text")
                    preview = next(child for child in window.winfo_children() if isinstance(child, tk.Toplevel))
                    self.assertEqual(source.read_text(encoding="utf-8"), "before")
                    preview.destroy()

                    class FakePreview:
                        destroyed = False

                        def destroy(self):
                            self.destroyed = True

                    fake_preview = FakePreview()
                    with patch.object(app.messagebox, "askyesno", return_value=True), \
                         patch.object(app.messagebox, "showinfo"):
                        window._replace_with_backup(source, "after", fake_preview, "text")
                    backups = list(Path(temp.name).glob("reviewed.txt.backup-*"))
                    self.assertEqual(source.read_text(encoding="utf-8"), "after")
                    self.assertEqual(len(backups), 1)
                    self.assertEqual(backups[0].read_text(encoding="utf-8"), "before")
                    self.assertTrue(fake_preview.destroyed)
                finally:
                    window.index.db.close()
                    window.destroy()
        finally:
            logger = logging.getLogger("local_file_manager")
            for handler in list(logger.handlers):
                if handler not in original_handlers:
                    logger.removeHandler(handler)
                    handler.close()
            app._APP_LOGGER = None


if __name__ == "__main__":
    unittest.main()
