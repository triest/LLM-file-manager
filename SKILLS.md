# App capabilities and operating notes

## File indexing

The user explicitly chooses the folders to index. The index is stored in `%APPDATA%\LocalFileManager`; it contains extracted text and file metadata, while source files stay in place. Large files and common system/build folders are skipped.

## Local model connection

The main window opens on the chat tab and keeps a local chat list. Chats can use normal, goal, or Cursor project mode. The working folder is per chat. Attachments and conversation content are persisted locally in `%APPDATA%\LocalFileManager\chats.json`.

The app connects to a local Ollama service at `127.0.0.1:11434` or LM Studio's OpenAI-compatible API, usually `http://127.0.0.1:1234/v1`. On startup it tries Ollama and may launch the installed service without opening a console. File contents are sent only to the selected localhost model when the user asks for an AI operation. Natural language file-finding requests query the local index directly; search reads only indexed folders.

## Organization and office files

Organization suggestions do not modify files. The user selects a file and invokes a move or rename action separately. No delete operation exists, and the model has no file-system tools. Text and source-code edits show a diff first; replacing the source requires explicit confirmation and a timestamped backup. DOCX edits are limited to selected paragraph text and retain the rest of the Office package. Excel files can receive validated, cell-addressed changes from a model plan; show old and new values before writing an output copy. Replacing XLSX/XLSM requires confirmation and a timestamped backup. XLS/XLSB output is converted to a new XLSX file.

## Checks

After changing the app, verify syntax and content search, then exercise the visible chat, saved chat list, goal/Cursor modes, source preview, attachments, save-as, and text/DOCX/Excel edit previews using temporary files. Launch through the desktop shortcut and confirm the local model returns a response. Review filesystem operations to ensure the model still has no way to delete files; do not use real user documents for smoke edits.
