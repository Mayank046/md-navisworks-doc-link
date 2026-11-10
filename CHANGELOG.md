# Changelog

## 1.0.1 — 2026-11-08

- No-admin, per-user installer: `MD_Navisworks_DocLink_NoAdmin_Setup_v1.0.1.exe`.
- Installs to `%APPDATA%\Autodesk\ApplicationPlugins\MD.MDDocLink.bundle`. No administrator rights required.
- Plugin assembly version remains 1.0.0 (installer packaging change only).

## 1.0.0 — 2026-11-08

Initial release.

- Link documents (PDF, folders, URLs) from an Excel list to matching Navisworks items.
- Key matching (exact / contains / starts with), preview, unmatched report, replace or append, MD-prefixed links only.
- Optional document-folder pattern, selection-set creation, open-linked-docs and remove-MD-links tools.
- Administrator setup EXE for Navisworks Manage and Simulate 2017–2027.
- Core unit tests for Excel I/O and matching (no Navisworks required to run them).
