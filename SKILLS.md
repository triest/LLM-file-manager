# App capabilities and operating notes

## Local model runtime

The app downloads the official llama.cpp Windows CUDA runtime and its CUDA DLL packages on first launch, then invokes its `llama-cli` process invisibly for each request. It does not connect to Ollama, LM Studio, a model server, or a cloud API. The one-time runtime download is about 750 MB. Installation output is written to `%APPDATA%\LocalFileManager\engine_install.log`.

Users can select a folder containing multiple GGUF files, Ollama-style manifest/blob storage, or add several individual GGUF files. The list is persisted in local config. One model is loaded per request, then the CLI exits and frees VRAM. The same model selector is available from Files, Chat, and Cursor tabs.

Every generation must explicitly select an available CUDA device and offload all model layers to the GPU; never silently fall back to CPU. The app verifies devices through llama.cpp --list-devices. The default context is 32,768 tokens with Q8 KV cache and Flash Attention. Local diagnostics are stored at %APPDATA%\LocalFileManager\app.log; logs may include model paths, timings, CUDA device details, and errors, but must never contain prompts or indexed file contents. Show a visible busy indicator during model loading/generation.

## File indexing and chat

The user explicitly chooses folders to index. The index is stored in `%APPDATA%\LocalFileManager`; it contains extracted text and metadata while source files remain in place. Chat history and attachments stay local. Natural-language file-finding requests query the local index; only indexed folders are searched.

The **Files and Search** tab includes its own chat pane. It shares the current chat list, selected local model, mode, working folder, and attachments with the other tabs. A selected search result is added as local prompt context; natural-language find requests query the local index directly. File edits must use the existing reviewed diff flow and explicit user approval.

Detect Russian file-search requests typed with an English keyboard layout and route them to the local index without sending chat history to the model. Keep request history within the model context window by dropping oldest turns; never silently truncate the current request or attached content. If the current file plus request cannot fit, show a clear explanation. Surface model errors in the transcript and write technical diagnostics to the local app log.

## Cursor workspace

Cursor is a three-pane main-window tab: project roots and files on the left, a multi-tab text/code editor in the center, and the active local chat on the right. One chat can include several project folders at once, such as a source project and a new destination project. The editor supports explicit save with a backup, unsaved-change prompts, line numbers, syntax highlighting, in-file find/replace, go-to-line, project content search, click-through definition navigation, and project-wide usage search for functions and variables. Local LLM code completions are previewed and explicitly accepted into the unsaved buffer before a separate save. PHP and Laravel projects receive language-aware prompts; PHP/Blade syntax is highlighted, Laravel is detected from `composer.json`, and `php -l` is an optional local syntax check when PHP CLI exists. No PHP runtime, Composer, model server, or cloud connection is required to edit. Each chat remembers project roots and conversation; users can switch chats. The active file and explicitly attached files provide context. AI document/code changes still show a review diff and require explicit confirmation before saving. The model cannot directly access the project filesystem or run code.

## Safety and office files

Never add a delete operation. Organization suggestions are proposals; moves and renames require user actions. Text and source edits show a diff first; replacing a source requires explicit confirmation and a timestamped backup. DOCX changes retain unrelated document content. Excel edits are validated cell-addressed operations and are reviewed before writing an output copy. XLS/XLSB outputs become a new XLSX file. Optional document parsers fail gracefully when unavailable.

## Checks

After changing the app, verify syntax and content search, then exercise visible chat from Files, Chat, and Cursor, saved chat list, goal mode, source preview, attachments, save-as, context trimming, and text/DOCX/Excel edit previews using temporary files. Launch through the desktop shortcut and confirm a real response from an explicitly selected CUDA device. Confirm errors and CUDA diagnostics are written to the local app log without prompt or document contents. Review filesystem operations to ensure the model cannot delete files; never use real user documents for smoke edits.
