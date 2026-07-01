#!/usr/bin/env python3
"""
LA Court E-Court Suite - native messaging host.

Bridges the Chrome extension's "Export" popup to the local desktop so the
Download button can kick off your existing Word mail-merge automation. When the
popup finishes writing the Order_*.xlsx into Downloads, it sends this host a
single message and this host launches the mail-merge template. Opening the
template creates a new document from it, which fires the template's
Document_New macro (RunMailMerge) - i.e. exactly what you do today by
double-clicking the .dotm, just triggered automatically.

This host does NOT reimplement the macro; it only launches the template. All
folder creation, PDF moving, summary doc, Outlook draft, and pdf_linker steps
continue to run inside your proven VBA.

Protocol (Chrome native messaging):
  * stdin/stdout carry length-prefixed JSON: a little-endian uint32 byte count
    followed by that many bytes of UTF-8 JSON.
  * We read one request, act, write one response, and exit.

Messages handled:
  {"action": "launchTemplate", "template": "<optional path override>"}
      -> {"ok": true, "method": "com"|"startfile", "template": "<path>"}
  {"action": "ping"} -> {"ok": true, "pong": true}

Setup: edit TEMPLATE_PATH below to point at your .dotm, then run install.ps1
(see README.md).
"""

import os
import sys
import json
import struct

# ---------------------------------------------------------------------------
# EDIT THIS: full path to your mail-merge template (.dotm).
# ---------------------------------------------------------------------------
TEMPLATE_PATH = (
    r"C:\Users\ZCoderre\OneDrive - Los Angeles Superior Court"
    r"\Desktop"
    r"\Mail Merge Order Template (Automatic).dotm"
)


def read_message():
    """Read one length-prefixed JSON message from stdin. Returns None on EOF."""
    raw_len = sys.stdin.buffer.read(4)
    if len(raw_len) < 4:
        return None
    msg_len = struct.unpack("<I", raw_len)[0]
    data = sys.stdin.buffer.read(msg_len)
    if len(data) < msg_len:
        return None
    return json.loads(data.decode("utf-8"))


def send_message(obj):
    """Write one length-prefixed JSON message to stdout."""
    encoded = json.dumps(obj).encode("utf-8")
    sys.stdout.buffer.write(struct.pack("<I", len(encoded)))
    sys.stdout.buffer.write(encoded)
    sys.stdout.buffer.flush()


def launch_template(path):
    """
    Open the template so its Document_New macro runs.

    Preferred: drive Word via COM and call Documents.Add(Template=...,
    NewTemplate=False), which reliably creates a new document FROM the template
    (firing Document_New). Falls back to os.startfile, whose default "New" verb
    for a template also creates a new document from it. Returns the method used.
    """
    try:
        import win32com.client  # part of pywin32; optional but most reliable
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = True
        word.Documents.Add(Template=path, NewTemplate=False)
        return "com"
    except Exception:
        # pywin32 missing or COM failed - shell-open the template instead.
        os.startfile(path)  # noqa: for Windows; default verb creates a new doc
        return "startfile"


def handle(msg):
    action = (msg or {}).get("action")

    if action == "ping":
        return {"ok": True, "pong": True}

    if action == "launchTemplate":
        path = (msg.get("template") or "").strip() or TEMPLATE_PATH
        if not os.path.exists(path):
            return {"ok": False, "error": "Template not found: " + path}
        method = launch_template(path)
        return {"ok": True, "method": method, "template": path}

    return {"ok": False, "error": "Unknown action: " + str(action)}


def main():
    try:
        msg = read_message()
        if msg is None:
            return
        send_message(handle(msg))
    except Exception as e:  # never crash silently; report back to the extension
        try:
            send_message({"ok": False, "error": str(e)})
        except Exception:
            pass


if __name__ == "__main__":
    # On Windows, force binary stdio so the length-prefix protocol isn't
    # mangled by newline translation.
    if sys.platform == "win32":
        import msvcrt
        msvcrt.setmode(sys.stdin.fileno(), os.O_BINARY)
        msvcrt.setmode(sys.stdout.fileno(), os.O_BINARY)
    main()
