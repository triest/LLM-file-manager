# App capabilities and operating notes

## File indexing

The user explicitly chooses the folders to index. The index is stored in `%APPDATA%\LocalFileManager`; it contains extracted text and file metadata, while source files stay in place. Large files and common system/build folders are skipped.

## Local model connection

The app connects to a local Ollama service at `127.0.0.1:11434`. On startup it tries to connect and may launch the installed Ollama service without opening a console. File contents are sent only to that local model endpoint when the user asks for an AI operation.

## Organization and office files

Organization suggestions do not modify files. The user selects a file and invokes a move or rename action separately. Office documents open through the Windows default application. DOCX, XLSX, PPTX and OpenDocument files are text-indexed from their local document containers.

## Checks

After changing the app, verify syntax, run a small content-indexing smoke check, and launch through the desktop shortcut. When a local model is available, make one short generation request to confirm the connection.
