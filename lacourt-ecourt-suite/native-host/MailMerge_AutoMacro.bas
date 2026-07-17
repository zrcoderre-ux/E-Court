'=============================================================
' MAIL MERGE AUTO MACRO  --  ThisDocument module
' File: MailMerge_AutoMacro.bas
'
' SETUP: Open .dotm in Word  >  Alt+F11  >  ThisDocument  >
'        select all existing code, delete it, paste this
'        entire file  >  Save
'
' EMAIL DELIVERY: Uses mailto: URI to open a New Outlook
' compose window with subject pre-filled. The compose window
' is then brought to the foreground, a real Ctrl+S keystroke
' is injected to save it to Drafts, and finally both the
' compose window and the main Outlook app window are
' minimized so neither is in your face. The draft is already
' saved - open it from Drafts (or restore the minimized
' compose window) when ready to attach the workup and send.
'=============================================================

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Window enumeration / show-state APIs for finding and
' minimizing the New Outlook compose window and main app
' window without relying on focus or message loops.
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
    (ByVal hwndParent As LongPtr, ByVal hwndChildAfter As LongPtr, _
     ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
Private Declare PtrSafe Function GetWindowTextA Lib "user32" _
    (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" _
    (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "user32" _
    (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function IsIconic Lib "user32" _
    (ByVal hwnd As LongPtr) As Long

' Focus + keystroke-injection APIs for saving the New Outlook
' compose draft. New Outlook is WebView2-hosted, so there is no
' COM object model for the compose window; the only reliable
' way to save the draft is to focus the window and inject a
' real Ctrl+S at the OS level (keybd_event goes to whichever
' window has keyboard focus, which is why the window MUST be
' foreground first - SendKeys alone is not reliable here).
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" _
    (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Sub keybd_event Lib "user32" _
    (ByVal bVk As Byte, ByVal bScan As Byte, _
     ByVal dwFlags As Long, ByVal dwExtraInfo As LongPtr)
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" _
    (ByVal hwnd As LongPtr, ByRef lpdwProcessId As Long) As Long
Private Declare PtrSafe Function AttachThreadInput Lib "user32" _
    (ByVal idAttach As Long, ByVal idAttachTo As Long, _
     ByVal fAttach As Long) As Long
Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long

Private Const SW_MINIMIZE As Long = 6
Private Const KEYEVENTF_KEYUP As Long = &H2
Private Const VK_CONTROL As Byte = &H11
Private Const VK_S As Byte = &H53

Private Sub Document_New()
    Call RunMailMerge
End Sub

'=============================================================
' MAIN ROUTINE
'=============================================================

Sub RunMailMerge()

    Dim oDoc            As Document
    Dim oMM             As MailMerge
    Dim oMergedDoc      As Document
    Dim sDownloadsPath  As String
    Dim sSummaryPath    As String
    Dim sExcelFile      As String
    Dim dtExcelTime     As Date
    Dim dtPrevExcelTime As Date
    Dim sFirstPara      As String
    Dim sSecondPara     As String
    Dim sFolderName     As String
    Dim sFolderPath     As String
    Dim sFileName       As String
    Dim sSavePath       As String
    Dim sSummaryName    As String
    Dim sSummarySavePath As String

    Set oDoc = ActiveDocument
    Set oMM = oDoc.MailMerge

    ' Suppress all Word dialogs/alerts for the duration of this macro
    Application.DisplayAlerts = wdAlertsNone

    '--- 1. Paths ---------------------------------------------
    sDownloadsPath = Environ("USERPROFILE") & "\Downloads\"
    sSummaryPath = "C:\Users\ZCoderre\OneDrive - Los Angeles Superior Court\Summaries\"

    '--- 2. Find most recent Order*.xlsx ----------------------
    sExcelFile = GetMostRecentExcel(sDownloadsPath, dtExcelTime, dtPrevExcelTime)

    If sExcelFile = "" Then
        MsgBox "No Excel file starting with 'Order' found in:" & vbCrLf & sDownloadsPath & _
               vbCrLf & vbCrLf & "Mail merge cancelled.", _
               vbExclamation, "Mail Merge - No File Found"
        Exit Sub
    End If

    '--- 3. Unlock fields (prevents "locked fields" dialog) ---
    Dim oField As Field
    For Each oField In oDoc.Fields
        oField.Locked = False
    Next oField

    '--- 4. Connect data source and execute merge -------------
    oMM.MainDocumentType = wdFormLetters
    On Error GoTo ErrHandler

    oMM.OpenDataSource _
        Name:=sExcelFile, _
        Format:=wdOpenFormatAuto, _
        Connection:="Provider=Microsoft.ACE.OLEDB.12.0;" & _
                    "Data Source=" & sExcelFile & ";" & _
                    "Extended Properties=""Excel 12.0 Xml;HDR=YES"";", _
        SQLStatement:="SELECT * FROM [Sheet1$]"

    oMM.Destination = wdSendToNewDocument
    oMM.Execute Pause:=False

    '--- 5. Get merged document -------------------------------
    Set oMergedDoc = ActiveDocument

    '--- 6. Read Para 1 (filename line) -----------------------
    sFirstPara = Trim(oMergedDoc.Paragraphs(1).Range.Text)
    If Right(sFirstPara, 1) = Chr(13) Or Right(sFirstPara, 1) = Chr(7) Then
        sFirstPara = Left(sFirstPara, Len(sFirstPara) - 1)
    End If
    sFirstPara = Trim(sFirstPara)

    '--- 7. Read Para 2 (email subject line) ------------------
    sSecondPara = Trim(oMergedDoc.Paragraphs(2).Range.Text)
    If Right(sSecondPara, 1) = Chr(13) Or Right(sSecondPara, 1) = Chr(7) Then
        sSecondPara = Left(sSecondPara, Len(sSecondPara) - 1)
    End If
    sSecondPara = Trim(sSecondPara)

    '--- 8. Build folder name (Case Number + Plaintiff) -------
    Dim iSpace As Long, iVs As Long
    iSpace = InStr(sFirstPara, " ")
    iVs = InStr(sFirstPara, " vs ")
    If iSpace > 0 And iVs > 0 Then
        Dim sCaseNum As String, sPlaintiff As String
        sCaseNum = Trim(Left(sFirstPara, iSpace - 1))
        sPlaintiff = Trim(Mid(sFirstPara, iSpace + 1, iVs - iSpace - 1))
        sFolderName = SanitizeName(sCaseNum & " " & sPlaintiff)
    Else
        sFolderName = SanitizeName(sFirstPara)
    End If
    sFolderPath = sDownloadsPath & sFolderName

    '--- 9. Build file name (safe, abbreviated, fits path) ----
    sFileName = BuildSafeFileName(sFirstPara, sFolderPath)

    '--- 10. Capture subject (Para 2) into a string variable --
    Dim sSubject    As String
    Dim oSubjRange  As Range
    Set oSubjRange = oMergedDoc.Paragraphs(2).Range
    oSubjRange.MoveEnd Unit:=wdCharacter, count:=-1
    sSubject = oSubjRange.Text

    '--- 11. Delete Para 1 and Para 2 in one range operation --
    Dim oDelRange As Range
    Set oDelRange = oMergedDoc.Range( _
        oMergedDoc.Paragraphs(1).Range.Start, _
        oMergedDoc.Paragraphs(3).Range.Start)
    oDelRange.Delete

    '--- 12. Clean up double spaces and space-before-comma ----
    Dim oFind As Find
    Set oFind = oMergedDoc.Content.Find
    With oFind
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = True
        .Text = " {2,}"
        .Replacement.Text = " "
        .Execute Replace:=wdReplaceAll
        .Text = "([A-Za-z]) ,"
        .Replacement.Text = "\1,"
        .Execute Replace:=wdReplaceAll
        .MatchWildcards = False
    End With

    '--- 13. Create Downloads folder --------------------------
    If sFolderName <> "" Then
        If Dir(sFolderPath, vbDirectory) = "" Then
            MkDir sFolderPath
        End If
    End If

    '--- 14. Save merged doc to Downloads folder --------------
    If sFileName = "" Then sFileName = sFolderName
    sSavePath = sFolderPath & "\" & sFileName & ".docx"
    oMergedDoc.SaveAs2 FileName:=sSavePath, FileFormat:=wdFormatXMLDocument

    '--- 15. Create blank Summary document --------------------
    sSummaryName = sFileName & " (Summary)"
    sSummarySavePath = sSummaryPath & sSummaryName & ".docx"

    On Error Resume Next  ' Summary creation is non-critical

    If Dir(sSummaryPath, vbDirectory) = "" Then
        MkDir sSummaryPath
    End If

    Dim oSummaryDoc As Document
    Set oSummaryDoc = Documents.Add(Template:="C:\Users\ZCoderre\OneDrive - Los Angeles Superior Court\Documents\Custom Office Templates\Summary Template.dotx", NewTemplate:=False)
    If Not oSummaryDoc Is Nothing Then
        oSummaryDoc.SaveAs2 FileName:=sSummarySavePath, FileFormat:=wdFormatXMLDocument
        oSummaryDoc.Close SaveChanges:=wdDoNotSaveChanges
        Set oSummaryDoc = Nothing
    End If

    On Error GoTo ErrHandler  ' Restore error handling

    '--- 16. Move PDFs ----------------------------------------
    MovePDFsToFolder sDownloadsPath, sFolderPath, dtExcelTime, dtPrevExcelTime

    '--- 17. Close merged doc ---------------------------------
    oMergedDoc.Close SaveChanges:=wdDoNotSaveChanges

    '--- 18. Open Outlook compose window via mailto: ----------
    '        Subject and full body are pre-filled. New Outlook
    '        opens a real compose window (not the .oft viewer)
    '        because mailto: is a registered URI handler.
    Call OpenOutlookDraft(sSubject)

    '--- 19. Run PDF linker -----------------------------------
    Call RunPDFLinker(sFolderPath)

    '--- 20. Open Downloads folder with new case folder selected -----
    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")
    oShell.Run "explorer.exe /select,""" & sFolderPath & """"

    '--- 21. Quit Word ----------------------------------------
    Application.Quit SaveChanges:=wdDoNotSaveChanges

    Exit Sub

ErrHandler:
    MsgBox "Mail merge failed." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "File used: " & sExcelFile, _
           vbCritical, "Mail Merge Error"
End Sub

'=============================================================
' OPEN NEW OUTLOOK DRAFT VIA mailto:, SAVE IT, MINIMIZE ALL
'
' Strategy:
'  1. Shell mailto:. New Outlook opens a compose window with
'     subject pre-filled.
'  2. Wait ~3 seconds for windows to appear and be enumerable.
'  3. Enumerate all top-level visible windows, identifying
'     - the compose window (title contains our subject)
'     - the main Outlook app window (title contains "Outlook"
'       but NOT our subject)
'  4. Save the draft: bring the compose window to the
'     foreground and inject a real Ctrl+S (keybd_event).
'     New Outlook saves the message to Drafts.
'  5. Minimize the compose window and the main app window.
'
' Why ShowWindow instead of WM_SYSCOMMAND/SC_MINIMIZE:
' New Outlook is a WebView2-hosted app and its compose window
' often ignores SC_MINIMIZE messages because the app handles
' window state through its own internal logic, not the
' standard Windows message loop. ShowWindow is a direct API
' call to the window manager - it bypasses the app's message
' handling and minimizes the window at the OS level.
'
' Why keybd_event instead of SendKeys for the Ctrl+S:
' SendKeys posts to the message queue of whatever VBA thinks
' is active, and Word is usually still the active app when
' this runs - the keystroke never reaches the WebView2
' compose surface. keybd_event synthesizes the keystroke at
' the OS input level, so as long as the compose window is
' genuinely foreground first, the Ctrl+S lands in it exactly
' as if typed - which is what saves the draft.
'=============================================================

Private Sub OpenOutlookDraft(sSubject As String)
    Dim sBody       As String
    Dim sMailto     As String
    Dim sRecipient  As String
    Dim oShell      As Object

    ' Recipient - set to "" to leave the To field blank
    sRecipient = "amackenzie@lacourt.ca.gov"

    ' Build body. vbCrLf = paragraph break in mailto.
    ' Outlook auto-appends the signature block, so it's omitted here.
    sBody = "Judge Mackenzie," & vbCrLf & vbCrLf & _
            "Please see the workup attached." & vbCrLf & vbCrLf & _
            "Best," & vbCrLf & _
            "Zach"

    ' Build mailto URI with URL-encoded subject and body.
    sMailto = "mailto:" & sRecipient & _
              "?subject=" & UrlEncode(sSubject) & _
              "&body=" & UrlEncode(sBody)

    Set oShell = CreateObject("WScript.Shell")
    On Error Resume Next
    oShell.Run sMailto, 1, False
    On Error GoTo 0
    Set oShell = Nothing

    ' Wait for the compose window and main Outlook window
    ' (if it wasn't already running) to appear.
    Sleep 3000

    ' Save the draft (Ctrl+S into the compose window), then
    ' minimize both the compose window and the main app window.
    Call SaveAndMinimizeOutlookWindows(sSubject)
End Sub

'=============================================================
' SAVE THE COMPOSE DRAFT, THEN MINIMIZE COMPOSE + MAIN WINDOW
'
' Pass 1 - FIND: walk all top-level visible windows and record
'  - the compose window (title contains our subject)
'  - main Outlook app windows (title contains "Outlook")
' Handles are only recorded here, never acted on: focusing or
' minimizing mid-walk reorders the z-order that FindWindowEx
' iterates, which can skip windows.
'
' Pass 2 - ACT:
'  - Compose window: ForceForeground + inject Ctrl+S
'    (SendCtrlS), give New Outlook a moment to commit the
'    save, then minimize it.
'  - Main Outlook window(s): minimize.
'
' Skips windows that are already minimized (IsIconic) so we
' don't waste calls or accidentally restore-then-minimize.
'
' Retries the find pass up to 3 times in case windows aren't
' yet enumerable on the first pass (Outlook startup can be
' slow if it wasn't already running).
'=============================================================

Private Sub SaveAndMinimizeOutlookWindows(sSubject As String)
    Dim hwnd            As LongPtr
    Dim sBuf            As String
    Dim nLen            As Long
    Dim sTitle          As String
    Dim sSubjMatch      As String
    Dim nAttempts       As Long
    Dim hwndCompose     As LongPtr
    Dim ahwndMain(7)    As LongPtr
    Dim nMainCount      As Long
    Dim i               As Long

    ' Use first 40 chars of subject as match key in case
    ' Outlook truncates the window title.
    sSubjMatch = Left(sSubject, 40)
    If Len(sSubjMatch) < 5 Then sSubjMatch = ""  ' too short to match safely

    For nAttempts = 1 To 3
        hwndCompose = 0
        nMainCount = 0

        '--- Pass 1: find (no side effects while walking) -----
        hwnd = 0
        Do
            hwnd = FindWindowEx(0, hwnd, vbNullString, vbNullString)
            If hwnd = 0 Then Exit Do
            If IsWindowVisible(hwnd) <> 0 And IsIconic(hwnd) = 0 Then
                sBuf = String(512, vbNullChar)
                nLen = GetWindowTextA(hwnd, sBuf, 512)
                If nLen > 0 Then
                    sTitle = Left(sBuf, nLen)

                    ' Compose window: title contains our subject
                    If sSubjMatch <> "" And _
                       InStr(1, sTitle, sSubjMatch, vbTextCompare) > 0 Then
                        hwndCompose = hwnd

                    ' Main Outlook app window: title contains "Outlook"
                    ' but not our subject
                    ElseIf InStr(1, sTitle, "Outlook", vbTextCompare) > 0 Then
                        If nMainCount <= UBound(ahwndMain) Then
                            ahwndMain(nMainCount) = hwnd
                            nMainCount = nMainCount + 1
                        End If
                    End If
                End If
            End If
        Loop

        ' Found the compose window (and whatever main windows
        ' exist) - act on them and stop retrying.
        If hwndCompose <> 0 Then Exit For

        ' Compose window not enumerable yet - wait and retry
        ' (especially likely on Outlook cold start).
        Sleep 1000
    Next nAttempts

    '--- Pass 2: act ------------------------------------------
    ' Save the draft FIRST, while the compose window is still
    ' restorable and before anything else steals focus.
    If hwndCompose <> 0 Then
        If ForceForeground(hwndCompose) Then
            ' Let the WebView2 compose surface finish wiring
            ' keyboard focus into the message body/fields.
            Sleep 500
            Call SendCtrlS
            ' Give New Outlook time to commit the save (it
            ' shows "Saved at ..." once done).
            Sleep 1000
        End If
        ShowWindow hwndCompose, SW_MINIMIZE
    End If

    For i = 0 To nMainCount - 1
        ShowWindow ahwndMain(i), SW_MINIMIZE
    Next i

    ' If the compose window was never found, nothing was saved
    ' or minimized - graceful degradation, the user saves and
    ' minimizes manually.
End Sub

'=============================================================
' FORCE A WINDOW TO THE FOREGROUND (so injected keys reach it)
'
' SetForegroundWindow is refused by Windows when the calling
' process isn't allowed to steal focus. Word usually IS the
' foreground app while this macro runs, so the plain call
' normally succeeds - but if it doesn't, retry with the
' AttachThreadInput trick: temporarily attach our input queue
' to the current foreground window's thread, which grants the
' focus-change right, then detach.
'
' Returns True only when the target window is verified to be
' the foreground window - never inject keystrokes otherwise,
' or the Ctrl+S lands in some other app.
'=============================================================

Private Function ForceForeground(hwndTarget As LongPtr) As Boolean
    Dim nTry        As Long
    Dim idCur       As Long
    Dim idFore      As Long
    Dim nPid        As Long

    For nTry = 1 To 5
        If GetForegroundWindow() = hwndTarget Then
            ForceForeground = True
            Exit Function
        End If

        SetForegroundWindow hwndTarget
        Sleep 200
        If GetForegroundWindow() = hwndTarget Then
            ForceForeground = True
            Exit Function
        End If

        ' AttachThreadInput fallback: borrow the foreground
        ' thread's focus-change privilege.
        idCur = GetCurrentThreadId()
        idFore = GetWindowThreadProcessId(GetForegroundWindow(), nPid)
        If idFore <> 0 And idFore <> idCur Then
            AttachThreadInput idCur, idFore, 1
            SetForegroundWindow hwndTarget
            AttachThreadInput idCur, idFore, 0
        End If
        Sleep 200
    Next nTry

    ForceForeground = (GetForegroundWindow() = hwndTarget)
End Function

'=============================================================
' INJECT A REAL Ctrl+S AT THE OS INPUT LEVEL
'
' keybd_event synthesizes hardware-level key events that go to
' the focused window, exactly like typing. Sequence: Ctrl down,
' S down, S up, Ctrl up, with small settle delays so the
' WebView2 input pipeline sees the modifier held during the S.
'=============================================================

Private Sub SendCtrlS()
    keybd_event VK_CONTROL, 0, 0, 0
    Sleep 50
    keybd_event VK_S, 0, 0, 0
    Sleep 50
    keybd_event VK_S, 0, KEYEVENTF_KEYUP, 0
    Sleep 50
    keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0
End Sub

'=============================================================
' URL-ENCODE A STRING FOR USE IN A mailto: URI
' Encodes reserved/unsafe characters per RFC 3986. Line
' breaks (vbCrLf) become %0D%0A which mailto handlers
' interpret as paragraph breaks.
'=============================================================

Private Function UrlEncode(sRaw As String) As String
    Dim i       As Long
    Dim sChar   As String
    Dim nCode   As Long
    Dim sResult As String

    For i = 1 To Len(sRaw)
        sChar = Mid(sRaw, i, 1)
        nCode = AscW(sChar)
        If nCode < 0 Then nCode = nCode + 65536  ' AscW returns signed Int

        Select Case nCode
            Case 48 To 57, 65 To 90, 97 To 122    ' 0-9, A-Z, a-z
                sResult = sResult & sChar
            Case 45, 46, 95, 126                  ' - . _ ~ (unreserved)
                sResult = sResult & sChar
            Case 32                               ' space -> %20
                sResult = sResult & "%20"
            Case 0 To 127                         ' other ASCII -> %XX
                sResult = sResult & "%" & Right("0" & Hex(nCode), 2)
            Case Else                             ' non-ASCII -> UTF-8 %XX%XX
                sResult = sResult & Utf8Encode(nCode)
        End Select
    Next i

    UrlEncode = sResult
End Function

'=============================================================
' UTF-8 ENCODE A SINGLE UNICODE CODEPOINT AS %XX%XX...
'=============================================================

Private Function Utf8Encode(nCode As Long) As String
    Dim s As String

    If nCode < &H80 Then
        s = "%" & Right("0" & Hex(nCode), 2)
    ElseIf nCode < &H800 Then
        s = "%" & Hex(&HC0 Or (nCode \ &H40)) & _
            "%" & Hex(&H80 Or (nCode And &H3F))
    Else
        s = "%" & Hex(&HE0 Or (nCode \ &H1000)) & _
            "%" & Hex(&H80 Or ((nCode \ &H40) And &H3F)) & _
            "%" & Hex(&H80 Or (nCode And &H3F))
    End If

    Utf8Encode = s
End Function

'=============================================================
' FIND THE MOST RECENT Order*.xlsx IN DOWNLOADS
'=============================================================

Private Function GetMostRecentExcel(sFolder As String, _
                                    ByRef dtExcelTime As Date, _
                                    ByRef dtPrevTime As Date) As String
    Dim sFile       As String
    Dim sFullPath   As String
    Dim dtThis      As Date
    Dim sBest       As String
    Dim dtBest      As Date
    Dim sPrev       As String
    Dim dtPrev      As Date

    dtBest = CDate("1/1/1970")
    dtPrev = CDate("1/1/1970")

    sFile = Dir(sFolder & "Order*.xlsx")
    Do While sFile <> ""
        sFullPath = sFolder & sFile
        dtThis = FileDateTime(sFullPath)
        If dtThis > dtBest Then
            dtPrev = dtBest
            sPrev = sBest
            dtBest = dtThis
            sBest = sFullPath
        ElseIf dtThis > dtPrev Then
            dtPrev = dtThis
            sPrev = sFullPath
        End If
        sFile = Dir()
    Loop

    GetMostRecentExcel = sBest
    dtExcelTime = dtBest
    dtPrevTime = dtPrev
End Function

'=============================================================
' BUILD A SAFE FILE NAME FROM THE FIRST PARAGRAPH
'
' Guarantees the returned name fits within the Windows path
' limit for the target folder, so SaveAs2 can never fail on a
' too-long name. The length budget is derived from the actual
' case-folder path (a deeper folder leaves less room), capped
' at 240 characters.
'
' Shortening strategy when the name is too long:
'   1. Abbreviate known motion types (MSJ / MSA).
'   2. If still too long, repeatedly trim one character from
'      the END of whichever segment is currently longest among
'      {plaintiff, defendant, motion type}, until it fits. The
'      case number is never trimmed.
'   3. Final hard truncate as a guarantee (only triggers if the
'      case number alone exceeds the budget).
'=============================================================

Private Function BuildSafeFileName(sRaw As String, sFolderPath As String) As String

    Dim nMax As Long
    ' Windows full-path budget ~255 chars (conservative margin
    ' under the 260 MAX_PATH limit). Reserve room for the
    ' trailing "\", the ".docx" extension, and the folder path.
    nMax = 255 - Len(sFolderPath) - Len("\") - Len(".docx")
    If nMax > 240 Then nMax = 240
    If nMax < 1 Then nMax = 1

    Dim s As String
    s = SanitizeName(Replace(sRaw, "/", "."))
    s = ReplaceMotionAbbrev(s)

    If Len(s) <= nMax Then
        BuildSafeFileName = s
        Exit Function
    End If

    '--- Parse into case number + trimmable segments ----------
    Dim sCaseNum    As String
    Dim sPlaintiff  As String
    Dim sDefendant  As String
    Dim sMotion     As String
    Dim bHasVs      As Boolean

    Dim iSpace As Long
    iSpace = InStr(s, " ")
    If iSpace = 0 Then
        ' No spaces at all - nothing to parse; hard-truncate.
        BuildSafeFileName = Left(s, nMax)
        Exit Function
    End If

    sCaseNum = Left(s, iSpace - 1)

    Dim sRest As String
    sRest = Mid(s, iSpace + 1)

    Dim iVs As Long
    iVs = InStr(1, sRest, " vs ", vbTextCompare)

    Dim sAfterVs As String
    If iVs > 0 Then
        bHasVs = True
        sPlaintiff = Trim(Left(sRest, iVs - 1))
        sAfterVs = Trim(Mid(sRest, iVs + 4))   ' " vs " is 4 chars
    Else
        bHasVs = False
        sPlaintiff = Trim(sRest)
        sAfterVs = ""
    End If

    ' Locate the motion type inside the segment that should hold
    ' it: after "vs" when present, otherwise the plaintiff blob.
    Dim sScan As String
    If bHasVs Then sScan = sAfterVs Else sScan = sPlaintiff

    Dim iMotion As Long
    iMotion = FindMotionStart(sScan)

    If iMotion > 0 Then
        sMotion = Trim(Mid(sScan, iMotion))
        If bHasVs Then
            sDefendant = Trim(Left(sScan, iMotion - 1))
        Else
            sPlaintiff = Trim(Left(sScan, iMotion - 1))
        End If
    Else
        sMotion = ""
        If bHasVs Then sDefendant = sScan
    End If

    '--- Trim the longest segment until the name fits ---------
    Dim sName As String
    sName = AssembleName(sCaseNum, sPlaintiff, sDefendant, sMotion, bHasVs)

    Do While Len(sName) > nMax
        Dim nP As Long, nD As Long, nM As Long
        nP = Len(sPlaintiff)
        nD = Len(sDefendant)
        nM = Len(sMotion)

        If nP = 0 And nD = 0 And nM = 0 Then Exit Do  ' nothing left to trim

        ' Cut one char from the END of the currently longest part.
        If nP >= nD And nP >= nM Then
            sPlaintiff = Left(sPlaintiff, nP - 1)
        ElseIf nD >= nP And nD >= nM Then
            sDefendant = Left(sDefendant, nD - 1)
        Else
            sMotion = Left(sMotion, nM - 1)
        End If

        sName = AssembleName(sCaseNum, sPlaintiff, sDefendant, sMotion, bHasVs)
    Loop

    ' Final hard guarantee (only if the case number alone overflows).
    If Len(sName) > nMax Then sName = Left(sName, nMax)

    sName = Trim(sName)
    ' Trailing separators can be left behind after trimming a part
    ' to empty; clean them up.
    Do While Len(sName) > 0 And (Right(sName, 1) = " " Or Right(sName, 3) = " vs")
        If Right(sName, 3) = " vs" Then
            sName = Left(sName, Len(sName) - 3)
        Else
            sName = Left(sName, Len(sName) - 1)
        End If
        sName = Trim(sName)
    Loop

    BuildSafeFileName = sName
End Function

'=============================================================
' REASSEMBLE THE FILE NAME FROM ITS PARTS
' Case number is the fixed prefix; the rest are trimmable.
'=============================================================

Private Function AssembleName(sCaseNum As String, _
                              sPlaintiff As String, _
                              sDefendant As String, _
                              sMotion As String, _
                              bHasVs As Boolean) As String
    Dim r As String
    r = sCaseNum
    If Len(sPlaintiff) > 0 Then r = r & " " & sPlaintiff
    If bHasVs And Len(sDefendant) > 0 Then r = r & " vs " & sDefendant
    If Len(sMotion) > 0 Then r = r & " " & sMotion
    AssembleName = Trim(r)
End Function

'=============================================================
' FIND THE 1-BASED START OF THE MOTION TYPE WITHIN A STRING
' Returns 0 if no motion keyword is found.
'=============================================================

Private Function FindMotionStart(sText As String) As Long
    Dim sUp As String
    sUp = UCase(sText)

    Dim aKeywords(9) As String
    aKeywords(0) = "MOTION"
    aKeywords(1) = "DEMURRER"
    aKeywords(2) = "MSJ"
    aKeywords(3) = "MSA"
    aKeywords(4) = "PETITION"
    aKeywords(5) = "APPLICATION"
    aKeywords(6) = "REQUEST"
    aKeywords(7) = "ORDER"
    aKeywords(8) = "OPPOSITION"
    aKeywords(9) = "REPLY"

    Dim iBest As Long
    iBest = 0

    Dim j As Integer
    For j = 0 To 9
        Dim iPos As Long
        iPos = InStr(1, sUp, aKeywords(j))
        If iPos > 0 Then
            If iBest = 0 Or iPos < iBest Then iBest = iPos
        End If
    Next j

    FindMotionStart = iBest
End Function

'=============================================================
' APPLY KNOWN ABBREVIATIONS
'=============================================================

Private Function ReplaceMotionAbbrev(s As String) As String
    Dim sUp As String
    sUp = UCase(s)

    Dim aFrom(1) As String, aTo(1) As String
    aFrom(0) = "MOTION FOR SUMMARY JUDGMENT"
    aTo(0) = "MSJ"
    aFrom(1) = "MOTION FOR SUMMARY ADJUDICATION"
    aTo(1) = "MSA"

    Dim i As Integer
    For i = 0 To 1
        Dim iPos As Long
        iPos = InStr(sUp, aFrom(i))
        If iPos > 0 Then
            s = Left(s, iPos - 1) & aTo(i) & Mid(s, iPos + Len(aFrom(i)))
            sUp = UCase(s)
        End If
    Next i

    ReplaceMotionAbbrev = s
End Function

'=============================================================
' STRIP CHARACTERS INVALID IN WINDOWS FOLDER/FILE NAMES
'=============================================================

Private Function SanitizeName(sRaw As String) As String
    Dim s As String
    s = sRaw

    If Right(s, 1) = Chr(13) Or Right(s, 1) = Chr(7) Then
        s = Left(s, Len(s) - 1)
    End If

    Dim aInvalid(8) As String
    aInvalid(0) = "\"
    aInvalid(1) = "/"
    aInvalid(2) = ":"
    aInvalid(3) = "*"
    aInvalid(4) = "?"
    aInvalid(5) = """"
    aInvalid(6) = "<"
    aInvalid(7) = ">"
    aInvalid(8) = "|"

    Dim i As Integer
    For i = 0 To 8
        s = Replace(s, aInvalid(i), "")
    Next i

    SanitizeName = Trim(s)
End Function

'=============================================================
' MOVE PDFs FROM DOWNLOADS INTO CASE FOLDER
'
' Only PDFs whose last-modified date is TODAY are eligible to
' be swept in. The candidate list is filtered to today's date
' at collection time, so both the main "newer than the order
' sheet" pass and the fallback band pass inherit the limit.
' Anything dated before today is ignored regardless of how its
' timestamp compares to the Excel file.
'=============================================================

Private Sub MovePDFsToFolder(sDownloadsPath As String, _
                              sFolderPath As String, _
                              dtCutoff As Date, _
                              dtPrevCutoff As Date)

    Dim aFiles()    As String
    Dim nCount      As Integer
    Dim sFile       As String
    Dim sFullPath   As String
    Dim sDestPath   As String
    Dim i           As Integer
    Dim dtToday     As Date

    ' Today's date with no time-of-day component. Int() on a
    ' FileDateTime strips the time portion so the comparison
    ' below is date-only (any PDF modified at any time today
    ' qualifies; anything from a prior day does not).
    dtToday = Date

    '--- Collect only PDFs last modified TODAY ----------------
    nCount = 0
    ReDim aFiles(0)
    sFile = Dir(sDownloadsPath & "*.pdf")
    Do While sFile <> ""
        sFullPath = sDownloadsPath & sFile
        If Int(FileDateTime(sFullPath)) = dtToday Then
            ReDim Preserve aFiles(nCount)
            aFiles(nCount) = sFile
            nCount = nCount + 1
        End If
        sFile = Dir()
    Loop

    '--- Main pass: today's PDFs newer than the order sheet ---
    Dim nMoved As Integer
    nMoved = 0
    For i = 0 To nCount - 1
        sFullPath = sDownloadsPath & aFiles(i)
        If FileDateTime(sFullPath) > dtCutoff Then
            sDestPath = sFolderPath & "\" & aFiles(i)
            If Dir(sDestPath) = "" Then
                On Error Resume Next
                Name sFullPath As sDestPath
                If Err.Number = 0 Then nMoved = nMoved + 1
                On Error GoTo 0
            End If
        End If
    Next i

    '--- Fallback: today's PDFs in (prevExcel, currentExcel] --
    If nMoved = 0 And dtPrevCutoff > CDate("1/1/1970") Then
        For i = 0 To nCount - 1
            sFullPath = sDownloadsPath & aFiles(i)
            Dim dtFile As Date
            dtFile = FileDateTime(sFullPath)
            If dtFile > dtPrevCutoff And dtFile <= dtCutoff Then
                sDestPath = sFolderPath & "\" & aFiles(i)
                If Dir(sDestPath) = "" Then
                    On Error Resume Next
                    Name sFullPath As sDestPath
                    On Error GoTo 0
                End If
            End If
        Next i
    End If

End Sub

'=============================================================
' RUN pdf_linker.py ON THE CASE FOLDER
'=============================================================

Private Sub RunPDFLinker(sFolderPath As String)
    Dim sScript As String
    Dim sCmd    As String
    Dim oShell  As Object

    sScript = "C:\Users\ZCoderre\Apps\PDF Linker\pdf_linker.py"

    If Dir(sScript) = "" Then Exit Sub

    sCmd = "pythonw """ & sScript & """ """ & sFolderPath & """ --provider lexis"

    Set oShell = CreateObject("WScript.Shell")

    ' Run non-blocking (False) so Word can quit immediately.
    ' The Python script continues in the background and finishes
    ' linking PDFs while the user reviews the Outlook draft.
    On Error Resume Next
    oShell.Run sCmd, 0, False
    On Error GoTo 0

    Set oShell = Nothing
End Sub


