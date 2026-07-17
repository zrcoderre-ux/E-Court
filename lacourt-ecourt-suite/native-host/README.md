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
| `MailMerge_AutoMacro.bas` | Source of record for the `.dotm`'s `ThisDocument` VBA. To update the template: open the `.dotm` in Word > Alt+F11 > `ThisDocument` > delete the existing code and paste this whole file > Save. |

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

## Agenda auto-advance on paste (optional)

The same host can also watch the OS clipboard so the **agenda page advances to
the next day automatically after you paste it into Excel** — no clicks, no
switching back to the browser.

How it works:

1. Landing on an agenda page auto-copies the cleaned agenda to your clipboard.
2. You paste it into Excel. Your paste macro then **clears the clipboard**.
3. The host (running in the background via a persistent connection) sees the
   clipboard go empty and tells the extension, which navigates to the next day
   (already prefetched, so it loads instantly).

### Enable it

1. **Update your Excel paste macro** so it clears the system clipboard at the
   end. `Module2.bas` in this repo is the ready-to-import version — it adds a
   `ClearSystemClipboard` helper (three `user32` API calls) and calls it in the
   macro's `CleanUp` block. `Application.CutCopyMode = False` alone does **not**
   clear the Windows clipboard, so the API call is required. It has no
   `Attribute` lines, so you can **File > Import File** it or paste its contents
   over your module. (Or edit your existing module: add the `#If VBA7 ... #End
   If` `Declare` block plus `ClearSystemClipboard` at the top, and call
   `ClearSystemClipboard` in `CleanUp`.) `Attribute ...` lines only work when a
   `.bas` is imported, not pasted -- that's the usual "syntax error on paste".
2. Open the extension's **Options** page and tick **“Auto-advance to the next
   agenda day after pasting.”** The setting persists (synced).
3. That's it — paste as usual and each paste jumps you to the next day.

Notes:

- It advances **one day per paste**, paced by you — it never runs ahead on its
  own. Untick the option any time to stop.
- On each advance it jumps to the next day that has a **green (will-be-copied)
  hearing**, skipping weekends, holidays, and fully-excluded days.
- When it advances, the service worker **refocuses the Chrome window** so the
  next page can auto-copy to the clipboard (Chrome blocks clipboard writes from
  an unfocused tab). Focus returns to eCourt; click back into Excel as usual.
- If you **scrub back** to an earlier day to review, auto-advance is suspended
  there (it only fires at/ahead of the furthest day you've reached), so pasting
  won't yank you forward.
- The cue is specifically the clipboard going *empty*; only your paste macro
  does that, so ordinary copying/pasting elsewhere won't trigger it.
- No `pywin32` needed for this — the clipboard is read via stdlib `ctypes`.

## Uninstall

```powershell
Remove-Item "HKCU:\Software\Google\Chrome\NativeMessagingHosts\com.lacourt.ecourt_host"
```
(and the Edge path if you registered it), then uncheck the auto-run box in the
popup — the Download button reverts to a plain spreadsheet save.
