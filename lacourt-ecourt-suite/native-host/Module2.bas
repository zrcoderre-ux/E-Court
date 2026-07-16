Option Explicit

' --- Windows API: clear the SYSTEM clipboard after pasting. This is the cue the
'     LA Court E-Court Suite extension watches for: when the clipboard goes empty
'     right after you paste the auto-copied agenda, it advances to the next day.
'     (Application.CutCopyMode = False only clears Excel's marching ants, not the
'     actual Windows clipboard, so we call the API directly.)
#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
#Else
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
#End If

Public Sub ClearSystemClipboard()
    ' Best-effort: if another app momentarily holds the clipboard, just skip.
    If OpenClipboard(0&) <> 0 Then
        EmptyClipboard
        CloseClipboard
    End If
End Sub

Sub PasteKeepHyperlinks()
    ' Pastes the clipboard into the active cell using the DESTINATION cells'
    ' existing formatting, while preserving any hyperlinks from the source.

    Dim wsDest As Worksheet, anchor As Range
    Dim wsTmp As Worksheet, srcRange As Range, target As Range
    Dim hl As Hyperlink, c As Range
    Dim r0 As Long, c0 As Long, rOff As Long, cOff As Long
    Dim savedVal As Variant
    Dim fName As String, fSize As Double
    Dim fBold As Boolean, fItalic As Boolean, fStrike As Boolean
    Dim fUnder As Long, fColor As Variant

    Set anchor = ActiveCell
    Set wsDest = anchor.Worksheet

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo CleanUp

    ' 1) Dump the clipboard onto a scratch sheet so we can read its hyperlinks
    Set wsTmp = wsDest.Parent.Worksheets.Add
    wsTmp.Paste Destination:=wsTmp.Range("A1")
    Set srcRange = wsTmp.UsedRange

    ' 2) Map onto the destination at the same dimensions
    Set target = anchor.Resize(srcRange.Rows.Count, srcRange.Columns.Count)

    ' 3) Move values only -> destination keeps all of its own formatting
    target.Value = srcRange.Value
    ' (swap .Value for .Formula on both sides if you need formulas preserved)

    ' 4) Re-create each hyperlink at the matching cell, then restore the
    '    destination font so the link doesn't repaint it blue/underlined
    r0 = srcRange.Cells(1, 1).Row
    c0 = srcRange.Cells(1, 1).Column
    For Each hl In srcRange.Hyperlinks
        rOff = hl.Range.Cells(1, 1).Row - r0
        cOff = hl.Range.Cells(1, 1).Column - c0
        Set c = target.Cells(1, 1).Offset(rOff, cOff)

        Dim addr As String, subAddr As String
        addr = hl.Address
        subAddr = hl.SubAddress

        savedVal = c.Value
        With c.Font
            fName = .Name: fSize = .Size
            fBold = .Bold: fItalic = .Italic
            fStrike = .Strikethrough: fUnder = .Underline
            fColor = .Color
        End With

        ' Only pass arguments that are non-empty
        On Error Resume Next
        If Len(addr) > 0 And Len(subAddr) > 0 Then
            wsDest.Hyperlinks.Add anchor:=c, Address:=addr, SubAddress:=subAddr
        ElseIf Len(addr) > 0 Then
            wsDest.Hyperlinks.Add anchor:=c, Address:=addr
        ElseIf Len(subAddr) > 0 Then
            wsDest.Hyperlinks.Add anchor:=c, Address:="", SubAddress:=subAddr
        End If
        On Error GoTo CleanUp

        c.Value = savedVal
        With c.Font
            .Name = fName: .Size = fSize
            .Bold = fBold: .Italic = fItalic
            .Strikethrough = fStrike: .Underline = fUnder
            .Color = fColor
        End With
    Next hl

CleanUp:
    If Not wsTmp Is Nothing Then wsTmp.Delete
    Application.CutCopyMode = False
    ' Clear the system clipboard so the extension knows the paste is done and can
    ' move to the next agenda day. (No-op for the merge; only matters when the
    ' agenda auto-advance toggle is on.)
    ClearSystemClipboard
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    wsDest.Activate
    anchor.Select
    If Err.Number <> 0 Then MsgBox "Paste failed: " & Err.Description, vbExclamation
End Sub
