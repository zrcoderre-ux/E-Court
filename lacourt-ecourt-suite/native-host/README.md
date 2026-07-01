# Native host — one-click Word mail merge from the Export popup

This lets the **Export** popup's **Download** button do exactly what you do
today by hand: after it writes the `Order_*.xlsx` into your Downloads folder,
it launches your Word mail-merge template, which fires the template's
`Document_New` macro (`RunMailMerge`) and runs the whole pipeline — folder
creation, PDF moves, summary doc, Outlook draft, `pdf_linker`, and Explorer.

The extension can't reach the desktop on its own (browser sandbox), so it talks
to this tiny local **native messaging host**, which launches the template. The
host does **not** reimplement your macro — your proven VBA stays in charge.

## Files

| File | What it is |
|---|---|
| `ecourt_host.py` | The host. Reads one message, launches your `.dotm`, replies, exits. |
| `ecourt_host.bat` | Launcher Chrome runs (calls `python ecourt_host.py`). |
| `com.lacourt.ecourt_host.json` | Host manifest (whitelists your extension). |
| `install.ps1` | Patches the manifest + registers the host in the registry. |

## Requirements

- Windows with Python installed. `ecourt_host.bat` calls the interpreter by an
  **absolute path** (Microsoft Store Python's bare `python` alias can fail when
  Chrome spawns the host). If your Python moves, update that path — find it with
  `python -c "import sys; print(sys.executable)"`.
- Recommended: `pywin32` (`pip install pywin32`) — lets the host start the
  template via Word COM, the most reliable way to fire `Document_New`. Without
  it the host falls back to `os.startfile`, which also works in most setups.

## Setup (once)

1. **Point the host at your template.** Open `ecourt_host.py` and set
   `TEMPLATE_PATH` to the full path of your `.dotm`
   (`Mail_Merge_Order_Template_Automatic.dotm`).

2. **Get your extension ID.** Go to `chrome://extensions`, enable *Developer
   mode*, and copy the ID under **LA Court E-Court Suite**.

3. **Register the host.** In PowerShell, from this `native-host` folder:

   ```powershell
   .\install.ps1 -ExtensionId <your-extension-id>
   ```

   (Add `-IncludeEdge` if you run the extension in Edge instead of Chrome.)

4. **Fully quit and reopen Chrome** so it picks up the new host.

## Test

- In the popup, leave **“Run the Word mail merge automatically after download”**
  checked and click **Download Spreadsheet**. The status line should read
  *“…— mail merge started.”* and Word should open a new merged document, after
  which your macro takes over.
- If it says *“Auto mail-merge did not run (…)”*, the spreadsheet still
  downloaded fine — only the host bridge didn't fire. See Troubleshooting.

## How the trigger stays correct

The macro finds its data with `Dir("Order*.xlsx")` in your Downloads folder and
picks the newest by timestamp. The popup saves as `Order_Template_Input_<case>.xlsx`
into the Downloads **root**, and waits for the download to finish writing
before messaging the host — so the macro always reads the sheet you just
exported.

## Troubleshooting

- **“Specified native messaging host not found.”** The registry key or manifest
  path is off, or Chrome wasn't restarted. Re-run `install.ps1` and restart
  Chrome completely.
- **“…host not set up” / no response.** Confirm `python` runs from a plain
  `cmd` prompt, and that `ecourt_host.bat`'s interpreter matches. You can smoke
  test the host directly: it should respond to a `{"action":"ping"}` message.
- **Word opens the template for editing instead of a new document.** Install
  `pywin32` so the host uses COM `Documents.Add` (which always creates a *new*
  document from the template).
- **Nothing lands in Downloads.** If you changed Chrome's download folder away
  from `%USERPROFILE%\Downloads`, the macro (which looks in `%USERPROFILE%\Downloads`)
  won't see the file. Keep them the same, or update the macro's `sDownloadsPath`.
- **A console window flashes on each Export.** Swap `python.exe` for
  `pythonw.exe` (same folder) in `ecourt_host.bat` (revert if the merge ever
  stops triggering).

## Uninstall

```powershell
Remove-Item "HKCU:\Software\Google\Chrome\NativeMessagingHosts\com.lacourt.ecourt_host"
```
(and the Edge path if you registered it), then uncheck the auto-run box in the
popup — the Download button reverts to a plain spreadsheet save.
