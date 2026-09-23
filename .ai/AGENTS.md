# Repository instructions

This repository contains Outlook VBA source snippets. Preserve compatibility with the classic Outlook VBA host; do not introduce dependencies on a standalone VB runtime.

- Treat the `outlook__*.vba` files as the source of truth. Do not edit files under ignored `backup/` archives.
- Keep `Option Explicit` in modules that already use it and prefer explicit Outlook object types where the existing module does so.
- Preserve the distinction between standard-module code and `ThisOutlookSession` event code. `outlook__ThisOutlookSession_DuplicateWatcher.vba` must remain suitable for merging into the special Outlook session module.
- `outlook__EmailCleanerAuto.vba` contains machine-specific `LOG_FILE_PATH` and `ARCHIVE_FOLDER_PATH` constants. Do not replace them with another user's paths or expose local data in documentation.
- Changes to duplicate matching, archive handling, retention, deletion, categories, or mailbox selection are behaviorally sensitive. Call out their effect explicitly.
- `frmProgress.frm` and `frmProgress.frx` provide the shared UserForm dependency for `outlook__EmailCleaner.vba` and `outlook__InboxSizeTools.vba`. Keep the pair together when updating or importing the form.
- There is no automated test or build harness. For verification, inspect the affected call paths, compile with `Debug > Compile VBAProject` in Outlook, and test against disposable messages or a non-production mailbox/profile.
- Do not claim Outlook runtime verification unless it was actually performed. Clearly distinguish static inspection from manual Outlook testing.
- Keep onboarding-managed documentation under `.ai/`. The root `AGENTS.md` is only a bootstrap that points here.
- Update `.ai/README.md`, `.ai/onboarding.md`, or `.ai/architecture.md` when installation steps, public macros, dependencies, or event flows change.
- For multi-step project work, maintain `.ai/tasks.md` at meaningful state changes: `[ ]` pending, `[~]` in progress, `[x]` completed, and `[!]` blocked. Add a short verification note to completed work. Do not churn the task list for simple questions or isolated read-only checks.
