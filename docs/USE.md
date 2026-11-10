# MD Doc Link v1.0.0 — how to use

Author: MD (Mayank Dholariya)

MD Doc Link reads an Excel list, such as a line list or an equipment list. It matches each row's key (line number or tag) to items in the Navisworks model and adds that row's documents (PDF, drawing, folder or web link) to those items as **Navisworks hyperlinks**.

It works with Navisworks Manage and Simulate 2017–2027 (64-bit). Freedom cannot load plugins, and you do not need Microsoft Excel.

> **Save the model.** Hyperlinks are stored in the open session. They are kept only after you save the model as **NWF** or **NWD**. If you close without saving, they are lost.

## Install
1. Close Navisworks.
2. Run `MD_Navisworks_DocLink_Setup_v1.0.0.exe` as administrator.
3. Tick the Navisworks versions to add it to, then click **Install**.
4. Open Navisworks and go to **Tool Add-ins** > **MD Doc Link**. The panel docks; you can drag it anywhere.

## Prepare the Excel file
- Use one row per key, with a **KEY** column whose values match a model property, e.g. `LINE No.` = `0-A-1001-150-PIP`.
- Put the documents in one or more **DOC** columns. Each cell can hold:
  - a file path, either absolute (`D:\Docs\ISO\0-A-1001.pdf`) or relative to the Excel file's folder (`ISO\0-A-1001.pdf`)
  - a folder (`D:\Docs\0-A-1001\`)
  - a web link (`https://…`)
  - a cell hyperlink: the link target is used
  - several documents separated by `;` or line breaks
- An optional **LINK NAME** column sets the link's name. Without one, the file name is used.
- The header row is detected automatically. If needed, set it by hand with **Header row**, where 0 means auto.

## Steps in the panel
1. **Excel**: choose the `.xlsx` and pick the sheet.
2. **Columns**: pick the KEY column and tick the DOC column(s). The panel guesses them by column name and by columns that contain hyperlinks. Optionally pick a LINK NAME column.
3. **Doc folder** (optional): to find files by key, enter a folder and a pattern such as `*{KEY}*.pdf`. `{KEY}` is replaced by each row's key.
   - **Include subfolders**: search below the folder as well.
   - **Use folder only when the row's document cells are empty**: default on.
   - **Loose file names**: ignores `- _ .` and spaces, so `0A1001.pdf` matches the key `0-A-1001`.
4. **Match**:
   - **Property**: the default is **Item > Name**. Click *Read properties from selected item* to choose any other category or property, e.g. a Plant 3D or E3D tag property.
   - **Mode**:
     - Exact
     - Contains key (whole part)
     - Contains key (anywhere)
     - Starts with key
   - Matching is always case-insensitive, and a leading `/` is ignored (E3D names).
   - **Ignore type word before '/'**: `PIPE /0-A-1` is matched as `0-A-1`.
   - **Ignore separators**: `- _ . /` and spaces are ignored.
   - **Link top-most match only**: the default. If a pipe and its branches all match, only the pipe gets the link.
   - **Search in**: Whole model, or Current selection (faster on large models).
5. **Preview**: shows the Excel row, key, documents, matched item count and status:
   - **Matched**: ready to link.
   - **Not found**: no model item has this key.
   - **Doc file missing**: the item was found but the file is not on disk.
   - **No document** or **No key**: the Excel cell is empty.

   Double-click a row to select its items in the model. Use **Export unmatched report…** to save the rows that need attention to Excel.
6. **Apply**:
   - **Replace**: removes earlier MD links on the item, then adds the new ones.
   - **Append**: keeps every existing link and adds the new ones. The same URL is never added twice.
   - **Skip documents not found on disk**: default on.
   - **Create selection sets**: makes folder **MD Doc Link** with the sets *MD Doc Link - Linked* and *MD Doc Link - Doc missing*.
   - Click **Add links to model**, then **save the NWF/NWD**.
7. **Tools**:
   - **Open linked docs**: select an item and open its links. The links can sit on the item or on a parent, and you get a list when there are several. You can also use the Navisworks *Links* display (View > Links, or right-click > Links).
   - **Remove MD links…**: removes only the links this tool added. It works on the selection, or on the whole model if nothing is selected.

## Good to know
- Links added by this tool are named **`MD | …`**. Hyperlinks from other sources are never removed.
- Press **Esc** to cancel a long preview, apply or remove.
- Links point to the document's location, so other users need the same drive or share path. Use UNC paths (`\\server\share\…`) for team models.
- Log file: `%APPDATA%\MDDocLink\mddoclink.log`. Send it to MD if anything goes wrong.
