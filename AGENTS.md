# Project guidance

- Before implementing each feature, use the `feature-branch-delivery` skill: create a dedicated branch, make a feature commit, check the requested behavior, push the branch, then merge it into `main` only after the check succeeds. Push the updated `main` after merging.
- Keep this project local-first: never send indexed file content to a remote service.
- The GUI is a native Tkinter application. Do not introduce a web server or require OpenServer.
- Keep core launch and file indexing on the Python standard library. Optional document parsers must fail gracefully.
- File changes must be explicit user actions. AI organization output is a proposal, not an automatic filesystem operation.
- Never add a delete operation. The model has no filesystem tools; text edits require a preview and explicit user confirmation, with a backup before replacing the source.
- In the Cursor editor, direct code changes are only written on explicit Save and must create a backup. LLM completions must be previewed and explicitly accepted into the unsaved editor buffer; they are written only after a separate explicit Save. Do not add file deletion.
- Keep Cursor tooling local: project symbol navigation and usages search run on-device and inspect only connected project roots. Do not send whole project trees to the model; only the active editor buffer and explicitly attached files may enter a prompt.
- Chat attachments and replies must stay on localhost. Saving model output requires a user-selected destination; project mode may expose only files explicitly attached to the model.
- The **Files and Search** tab includes an integrated chat. File-search requests must use the local index only; the selected result may be added as prompt context on-device. Any file change requested there must use the reviewed diff and explicit user confirmation, with a backup before replacing a source. Keep the chat list, model, attachments, and working folder in sync with the other chat surfaces.
- The desktop launcher is created with `CreateShortcut.ps1`; the app entry point is `app.py`.
- When adding features, update `README.md` and check the shortcut launch, file indexing/search, local model connection, and reviewed text/DOCX editing. Use temporary samples for edit checks.
- Local model generations in every tab must explicitly use an available CUDA device; do not silently fall back to CPU. Show a visible busy indicator and record runtime errors/CUDA diagnostics in %APPDATA%\LocalFileManager\app.log, without recording prompts or indexed content. Bound old chat history to the context size, preserve the current request and attachments, and explain when they exceed the safe context limit. Test Files, Chat, and Cursor send flows plus keyboard-layout-aware local file search.
