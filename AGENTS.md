# Project guidance

- Before implementing each feature, use the `feature-branch-delivery` skill: create a dedicated branch, make a feature commit, check the requested behavior, push the branch, then merge it into `main` only after the check succeeds. Push the updated `main` after merging.
- Keep this project local-first: never send indexed file content to a remote service.
- The GUI is a native Tkinter application. Do not introduce a web server or require OpenServer.
- Keep core launch and file indexing on the Python standard library. Optional document parsers must fail gracefully.
- File changes must be explicit user actions. AI organization output is a proposal, not an automatic filesystem operation.
- Never add a delete operation. The model has no filesystem tools; text edits require a preview and explicit user confirmation, with a backup before replacing the source.
- Chat attachments and replies must stay on localhost. Saving model output requires a user-selected destination; project mode may expose only files explicitly attached to the model.
- The desktop launcher is created with `CreateShortcut.ps1`; the app entry point is `app.py`.
- When adding features, update `README.md` and check the shortcut launch, file indexing/search, local model connection, and reviewed text/DOCX editing. Use temporary samples for edit checks.
