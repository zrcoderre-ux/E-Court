#!/usr/bin/env python3
"""
LA Court E-Court Suite - native messaging host.

Two jobs:

1. Mail-merge trigger (one-shot, via chrome.runtime.sendNativeMessage). When the
   Export popup finishes writing Order_*.xlsx into Downloads it sends one
   {"action":"launchTemplate"} message; we open the Word template so its
   Document_New macro (RunMailMerge) fires. Opening the template creates a new
   document from it - exactly what double-clicking the .dotm does, automated.

2. Clipboard watch (persistent, via chrome.runtime.connectNative). The agenda
   page auto-copies the cleaned agenda to the clipboard on landing. When you
   paste it into Excel, your paste macro clears the clipboard. This host polls
   the OS clipboard in the background (no browser focus needed) and, when it
   sees the clipboard go from non-empty to empty, sends {"event":"clipboardEmpty"}
   so the extension can advance to the next agenda day. A periodic
   {"event":"keepalive"} keeps the extension's MV3 service worker awake.

Protocol (Chrome native messaging): stdin/stdout carry length-prefixed JSON - a
little-endian uint32 byte count then that many bytes of UTF-8 JSON. We loop
reading requests until stdin closes (EOF), acting on each. sendNativeMessage
closes the pipe after our one reply, so the one-shot mail-merge path still
exits cleanly.

Messages handled:
  {"action": "launchTemplate", "template": "<optional path override>"}
      -> {"ok": true, "method": "com"|"startfile", "template": "<path>"}
  {"action": "watchClipboard"}      -> {"ok": true, "watching": true}
  {"action": "stopClipboardWatch"}  -> {"ok": true, "watching": false}
  {"action": "ping"}                -> {"ok": true, "pong": true}

Setup: edit TEMPLATE_PATH below to point at your .dotm, then run install.ps1
(see README.md).
"""

import os
import sys
import json
import struct
import threading
import time

# ---------------------------------------------------------------------------
# EDIT THIS: full path to your mail-merge template (.dotm).
# ---------------------------------------------------------------------------
TEMPLATE_PATH = r"C:\Users\ZCoderre\Mail Merge Order Template (Automatic).dotm"

# stdout is written from both the main thread (responses) and the clipboard
# watch thread (events), so guard it.
_stdout_lock = threading.Lock()


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
    """Write one length-prefixed JSON message to stdout (thread-safe)."""
    encoded = json.dumps(obj).encode("utf-8")
    with _stdout_lock:
        sys.stdout.buffer.write(struct.pack("<I", len(encoded)))
        sys.stdout.buffer.write(encoded)
        sys.stdout.buffer.flush()


# ---------------------------------------------------------------------------
# Clipboard reading (Windows, via ctypes - no pywin32 dependency)
# ---------------------------------------------------------------------------

def get_clipboard_text():
    """
    Return the clipboard's unicode text, "" when the clipboard holds no text
    (e.g. it was emptied), or None when the clipboard couldn't be opened right
    now (another process holds it - retry next tick).

    ctypes restypes are set explicitly so 64-bit HANDLEs aren't truncated to
    32-bit ints (a classic ctypes footgun that yields a bad pointer).
    """
    try:
        import ctypes
        from ctypes import wintypes
    except Exception:
        return None
    CF_UNICODETEXT = 13
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32

    user32.OpenClipboard.argtypes = [wintypes.HWND]
    user32.OpenClipboard.restype = wintypes.BOOL
    user32.IsClipboardFormatAvailable.argtypes = [wintypes.UINT]
    user32.IsClipboardFormatAvailable.restype = wintypes.BOOL
    user32.GetClipboardData.argtypes = [wintypes.UINT]
    user32.GetClipboardData.restype = wintypes.HANDLE
    kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalLock.restype = ctypes.c_void_p
    kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]

    if not user32.OpenClipboard(None):
        return None
    try:
        if not user32.IsClipboardFormatAvailable(CF_UNICODETEXT):
            return ""
        handle = user32.GetClipboardData(CF_UNICODETEXT)
        if not handle:
            return ""
        ptr = kernel32.GlobalLock(handle)
        if not ptr:
            return ""
        try:
            text = ctypes.c_wchar_p(ptr).value
        finally:
            kernel32.GlobalUnlock(handle)
        return text or ""
    finally:
        user32.CloseClipboard()


# ---------------------------------------------------------------------------
# Clipboard watch thread
# ---------------------------------------------------------------------------

_watch_stop = threading.Event()
_watch_thread = None

# How often to poll, and how often to nudge the service worker so it doesn't
# idle out (MV3 terminates an idle worker after ~30s; any port message resets
# that timer).
POLL_SECONDS = 0.4
KEEPALIVE_SECONDS = 20


def _watch_loop():
    # `armed` becomes True once we've seen text on the clipboard (our agenda
    # copy). The paste macro then empties it, and that non-empty -> empty
    # transition is the cue. Re-arms on the next copy so every page works.
    armed = False
    last_keepalive = time.time()
    while not _watch_stop.is_set():
        text = get_clipboard_text()
        if text is not None:
            if len(text) > 0:
                armed = True
            elif armed:
                armed = False
                send_message({"event": "clipboardEmpty"})
        now = time.time()
        if now - last_keepalive >= KEEPALIVE_SECONDS:
            last_keepalive = now
            send_message({"event": "keepalive"})
        _watch_stop.wait(POLL_SECONDS)


def start_watch():
    global _watch_thread
    if _watch_thread and _watch_thread.is_alive():
        return
    _watch_stop.clear()
    _watch_thread = threading.Thread(target=_watch_loop, daemon=True)
    _watch_thread.start()


def stop_watch():
    _watch_stop.set()


# ---------------------------------------------------------------------------
# Mail-merge template launch
# ---------------------------------------------------------------------------

def launch_template(path):
    """
    Open the template so its Document_New macro runs. Prefer Word COM
    (Documents.Add(Template=..., NewTemplate=False), which reliably creates a
    new doc FROM the template); fall back to os.startfile, whose default "New"
    verb for a template also creates a new document from it.
    """
    try:
        import win32com.client  # pywin32; optional but most reliable
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = True
        word.Documents.Add(Template=path, NewTemplate=False)
        return "com"
    except Exception:
        os.startfile(path)  # noqa: Windows; default verb creates a new doc
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

    if action == "watchClipboard":
        start_watch()
        return {"ok": True, "watching": True}

    if action == "stopClipboardWatch":
        stop_watch()
        return {"ok": True, "watching": False}

    return {"ok": False, "error": "Unknown action: " + str(action)}


def main():
    # Loop so a persistent connectNative port can issue several commands and
    # receive events. sendNativeMessage closes stdin after our single reply, so
    # the one-shot mail-merge path breaks out of this loop and exits.
    while True:
        try:
            msg = read_message()
        except Exception as e:
            try:
                send_message({"ok": False, "error": str(e)})
            except Exception:
                pass
            break
        if msg is None:
            break
        try:
            resp = handle(msg)
            if resp is not None:
                send_message(resp)
        except Exception as e:
            try:
                send_message({"ok": False, "error": str(e)})
            except Exception:
                pass
    stop_watch()


if __name__ == "__main__":
    # On Windows, force binary stdio so the length-prefix protocol isn't
    # mangled by newline translation.
    if sys.platform == "win32":
        import msvcrt
        msvcrt.setmode(sys.stdin.fileno(), os.O_BINARY)
        msvcrt.setmode(sys.stdout.fileno(), os.O_BINARY)
    main()
