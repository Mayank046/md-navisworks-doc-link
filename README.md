# MD Doc Link

A dockable panel for Navisworks that links documents (PDF, folders, URLs) to model items directly from an Excel list — for example a line list or an equipment list. It matches each row's key (line number or tag) to items in the model and adds the row's documents as Navisworks hyperlinks on those items. Microsoft Excel does not need to be installed; `.xlsx` is read directly.

**Author:** MD (Mayank Dholariya)  
**Version:** 1.0.1 (no-admin installer; plugin assembly 1.0.0)  
**License:** MIT

## Screenshots

_Add screenshots of the dock pane, the Preview grid (Matched / Not found / Doc file missing), and a linked item in Navisworks in `docs/images/`._

## Supported Navisworks versions

Autodesk Navisworks **Manage** and **Simulate** **2017–2027** (64-bit). Navisworks Freedom cannot load plugins and is not supported.

## Install (no administrator rights)

Download the latest installer from [Releases](../../releases). Close Navisworks first.

- **Recommended (v1.0.1):** run `MD_Navisworks_DocLink_NoAdmin_Setup_v1.0.1.exe`. It installs for your Windows user only, to `%APPDATA%\Autodesk\ApplicationPlugins\MD.MDDocLink.bundle`. No admin rights, no Program Files.
- **Uninstall:** Windows Settings → Apps → "MD Doc Link".

v1.0.0 shipped an administrator setup (`MD_Navisworks_DocLink_Setup_v1.0.0.exe`) that installed under Program Files. Use the v1.0.1 no-admin installer on new machines.

## Use

1. Open the NWD / NWF. **Tool add-ins → MD Doc Link**.
2. Browse to an `.xlsx`, pick the sheet (header row 0 = auto-detect).
3. Choose the **KEY** column (for example LINE No. / TAG) and tick one or more **DOC** columns (file path, folder or URL; several per cell, split by `;` or a new line). An optional **LINK NAME** column is supported.
4. Optionally set a document folder + pattern such as `*{KEY}*.pdf`.
5. Preview: check Matched / Not found / Doc file missing. Double-click a row to select its items. Export an unmatched report if needed.
6. Apply links (**Replace** MD links or **Append**), then **save the NWF/NWD**.

Links added by this tool are named `MD | …`. Hyperlinks from other sources are never removed. The log is `%APPDATA%\MDDocLink\mddoclink.log`.

The full option list is in [docs/USE.md](docs/USE.md) and `installer/README.txt`.

## Build from source

You need the .NET SDK (8 or later is fine for tests; plugin builds target .NET Framework 4.6–4.8).

**Do not commit Autodesk Navisworks DLLs.** They are not redistributable. Point `NwRefDir` at your own Navisworks install (this plugin also needs `Autodesk.Navisworks.ComApi.dll` and `Autodesk.Navisworks.Interop.ComApi.dll` from the same folder).

```bash
# Core unit tests (no Navisworks required)
dotnet test tests/DocLink.Tests/DocLink.Tests.csproj -c Release

# Plugin DLL for one year, against your Navisworks install
dotnet build src/MDDocLink/MDDocLink.csproj -c Release \
  -p:NwYear=2024 \
  -p:NwRefDir="C:/Program Files/Autodesk/Navisworks Manage 2024" \
  -p:NwTfm=net48
```

`NwTfm` is `net46` for 2017–2018, `net47` for 2019–2020, and `net48` for 2021–2027.

To build every year, set `NW_REFS` to a directory of per-year API folders (`$NW_REFS/2017`, …) and run `./build_all.sh`. That script always runs the Core tests first.

Compile the no-admin installer with [Inno Setup 6](https://jrsoftware.org/isinfo.php) from `installer/DocLink_NoAdmin.iss` after staging the bundle to `stage/MD.MDDocLink.bundle/` (one DLL per year under `Contents/<year>/`, plus `PackageContents.xml`). The older administrator script is `installer/DocLink.iss`.

## Tests

```bash
dotnet test tests/DocLink.Tests/DocLink.Tests.csproj -c Release
```

The Core tests build temporary sample workbooks and do not need Navisworks.

## License

MIT © 2026 MD (Mayank Dholariya). See [LICENSE](LICENSE).
