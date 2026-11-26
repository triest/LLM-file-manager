# App capabilities and operating notes

## Local model runtime

The app downloads the official llama.cpp Windows CUDA runtime and its CUDA DLL packages on first launch, then invokes its `llama-cli` process invisibly for each request. It does not connect to Ollama, LM Studio, a model server, or a cloud API. The one-time runtime download is about 750 MB. Installation output is written to `%APPDATA%\LocalFileManager\engine_install.log`.

Users can select a folder containing multiple GGUF files, Ollama-style manifest/blob storage, or add several individual GGUF files. The list is persisted in local config. One model is loaded per request, then the CLI exits and frees VRAM. The same model selector is available from Files, Chat, and Cursor tabs.

## File indexing and chat

The user explicitly chooses folders to index. The index is stored in `%APPDATA%\LocalFileManager`; it contains extracted text and metadata while source files remain in place. Chat history and attachments stay local. Natural-language file-finding requests query the local index; only indexed folders are searched.

## Cursor workspace

Cursor is a three-pane main-window tab: project roots and files on the left, the selected file preview in the center, and the active local chat on the right. One chat can include several project folders at once, such as a source project and a new destination project. Each chat remembers its own project roots and conversation; users can switch chats without losing their workspace. The active file and explicitly attached files provide context. Text/document changes show a review preview and require explicit confirmation before saving. The model cannot directly access the project filesystem or run code.

## Safety and office files

Never add a delete operation. Organization suggestions are proposals; moves and renames require user actions. Text and source edits show a diff first; replacing a source requires explicit confirmation and a timestamped backup. DOCX changes retain unrelated document content. Excel edits are validated cell-addressed operations and are reviewed before writing an output copy. XLS/XLSB outputs become a new XLSX file. Optional document parsers fail gracefully when unavailable.

## Checks

After changing the app, verify syntax and content search, then exercise the visible chat, saved chat list, goal/Cursor modes, source preview, attachments, save-as, and text/DOCX/Excel edit previews using temporary files. Launch through the desktop shortcut and confirm a local model response. Review filesystem operations to ensure the model cannot delete files; never use real user documents for smoke edits.
