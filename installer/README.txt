MD Doc Link v1.0.1  -  link documents to Navisworks items directly from Excel
Author / publisher: MD (Mayank Dholariya)
Copyright (c) 2026 MD (Mayank Dholariya)

NEEDS
  Autodesk Navisworks Manage or Simulate, 2017 to 2027, 64-bit.
  (Navisworks Freedom cannot run plugins.)
  Microsoft Excel is NOT needed (.xlsx is read directly).

INSTALL
  1. Close Navisworks.
  2. Run MD_Navisworks_DocLink_NoAdmin_Setup_v1.0.1.exe (no administrator rights).
  3. Next > Install. The add-in goes into your user profile and Navisworks 2017-2027 loads it automatically.
  Uninstall: Windows Settings > Apps > "MD Doc Link".

USE (short)
  1. Open the NWD / NWF model. Tool Add-ins tab > MD Doc Link.
  2. Excel: browse to the .xlsx, pick the sheet (header row 0 = auto-detect).
  3. Columns: KEY column (e.g. LINE No. / TAG), tick DOC column(s)
     (cell = file path, folder or URL; several per cell split by ; or new line),
     optional LINK NAME column.
  4. Optional doc folder + pattern, e.g.  *{KEY}*.pdf  (finds files by key).
  5. Match: property (default Item > Name), mode (Exact / Contains / Starts with),
     scope Whole model or Current selection.
  6. Preview: check Matched / Not found / Doc file missing.
     Double-click a row to select its items. Export unmatched report if needed.
  7. Apply links (Replace MD links or Append).
  8. SAVE the model as NWF or NWD - links are kept only after you save.

  Click an item > Tools > Open linked docs, or use Navisworks' own Links.
  Remove MD links removes only links this tool added ("MD | " names).

NOTES
  - Full guide: USE.md
  - Log: %APPDATA%\MDDocLink\mddoclink.log   (send it to MD if anything fails)
