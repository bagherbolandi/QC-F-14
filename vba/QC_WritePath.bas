Attribute VB_Name = "QC_WritePath"
Option Explicit
'=======================================================================
'  QC-F-14  v2.0  --  WRITE PATH ONLY  (import this .bas, nothing else)
'-----------------------------------------------------------------------
'  ARCHITECTURE RULE (agreed in phase 1 of the redesign):
'    * formulas compute, VBA writes. This module never calculates a KPI,
'      never refreshes a pivot, never edits history and never deletes a QC
'      record. It writes exactly: one new tblQC row, one LOG_AUDIT line,
'      and the ROW_STATUS flag that the DQ rules already define.
'    * QC_VoidRow is the one non-append write allowed; QC_Import_Staging clears
'      the scratch row it just consumed (the STAGING sheet's own contract).
'    * every business rule is read from the workbook tables
'      (RULE_STATION / DIM_DISPOSITION) and every user-visible string from
'      Settings!QC_MSG -- so the rule lives on the sheet, not in code, and
'      changing a rule does not need a code review.
'
'  WHY THIS SOURCE IS ASCII-ONLY:
'    The VBE is an ANSI editor: Persian typed into a .bas gets re-encoded
'    through the Windows codepage on Import (and ZWNJ / U+06CC are not even
'    representable in cp1256). So there is not one Persian literal here.
'    Status strings come from Helper!CODE_ROW_STATUS, the defect/cavity
'    verdicts from Helper!CODE_MOLD_CHECK, event types from the Master
'    tables, labels/messages from Settings!QC_MSG. All of them arrive as
'    real Unicode from the file -- no codepage can mangle them.
'
'  CALENDAR: the Jalali -> Gregorian bridge is a lookup in the Helper
'    calendar grid (the very same grid the in-cell formulas use), so the
'    algorithm is NOT duplicated in VBA and cannot drift from it.
'
'  BARCODE: treated as a 19-character string, never as a number. The cell
'    is formatted Text before the value is assigned, so leading zeros
'    survive. Duplicate barcodes are NOT rejected: the row is written with
'    ROW_STATUS = "suspect duplicate" (item 4 of CODE_ROW_STATUS) so the
'    history stays intact and DQ shows it.
'
'  KNOWN LIMITS (do not pretend otherwise):
'    * LOG_AUDIT is best-effort: a user can edit this sheet. It is not a
'      legal audit trail and is not presented as one.
'    * duplicate detection uses COUNTIF on a 19-digit text key; COUNTIF
'      compares long digit strings as text, which is what we want here.
'    * one file per station. Excel has no record locking: two people
'      appending to the same file can lose a row. Aggregation is central.
'
'  ENTRY POINTS (assign to buttons / shortcuts on the FORM sheet):
'    QC_Append          FORM  -> one row into tblQC  (+ LOG_AUDIT)
'    QC_Import_Staging  STAGING rows flagged "ready" -> tblQC, then cleared
'    QC_VoidRow         selected tblQC row -> ROW_STATUS = void (no delete)
'    QC_ClearForm       clears the barcode only (next scan is ready)
'    QC_About           what this is, version, honest limits
'=======================================================================

Private Const TBL_QC As String = "tblQC"
Private Const TBL_RULE As String = "RULE_STATION"
Private Const TBL_DISP As String = "DIM_DISPOSITION"
Private Const GRID_MOLD As String = "CODE_MOLD_CHECK"
Private Const GRID_STAT As String = "CODE_ROW_STATUS"
Private Const CFG_SHEETS As String = "QC_SHEETS"
Private Const CFG_MSG As String = "QC_MSG"
Private Const BC_LEN As Long = 19
Private Const LOG_FIRST As Long = 3      ' LOG_AUDIT header is row 2
Private Const LOG_LAST As Long = 202    ' 200 events -- ring, never overwritten
' Column-index contracts of the Helper grids. tools/test_workbook.py T-W12
' fails if the workbook ever changes the order of these two lists.
Private Const MC_MATCH As Long = 1       ' code: cavity letter == barcode digit 4
Private Const MC_MISS As Long = 2        ' code: cavity letter <> digit 4
Private Const MC_DIRECTIONAL As Long = 3 ' mold is right/left, no cavity to test
Private Const MC_DIGIT As Long = 4       ' barcode digit 4 outside 1..5
Private Const ST_OK As Long = 1          ' confirmed
Private Const ST_VOID As Long = 3        ' voided
Private Const ST_DUP As Long = 4         ' suspect duplicate
Public Const QC_VERSION As String = "2.0.0-writepath"


'-----------------------------------------------------------------------
' 1. QC_Append -- the only way a record enters tblQC from the shop floor
'-----------------------------------------------------------------------
Public Sub QC_Append()
    Dim wsD As Worksheet, wsF As Worksheet
    Dim pick(1 To 9) As String, code(1 To 9) As String, disp(1 To 9) As String
    Dim i As Long, r As Long, bc As Long, msg As String
    Dim et As String, status As String

    On Error GoTo Fail
    Set wsF = SheetAt("FORM")
    Set wsD = SheetAt("DATA")

    ' (a) the nine input cells -- the same cells the data validation drives
    pick(1) = CellTxt(wsF, "C6")    ' Jalali event date, from the form's list
    pick(2) = CellTxt(wsF, "C8")    ' shift
    pick(3) = CellTxt(wsF, "C10")   ' station
    pick(4) = CellTxt(wsF, "C12")   ' inspector
    pick(5) = CellTxt(wsF, "F6")    ' part
    pick(6) = CellTxt(wsF, "F8")    ' mold / cavity
    pick(7) = CellTxt(wsF, "F10")   ' defect
    pick(8) = CellTxt(wsF, "F12")   ' disposition (scrap / rework / return)
    bc = 9
    pick(bc) = RawText(wsF, "C14")  ' barcode -- deliberately not trimmed-away-safe

    For i = 2 To 8
        code(i) = PickCode(pick(i))
        disp(i) = PickName(pick(i))
        If Len(code(i)) = 0 Then FailNow Msg("E_INCOMPLETE") & vbCrLf _
            & CellTxt(wsF, "C28")
    Next i

    ' (b) barcode: exactly 19 digits, and it must still be TEXT
    If Len(pick(bc)) <> BC_LEN Or Not AllDigits(pick(bc)) Then FailNow Msg("E_BARCODE")

    ' (c) calendar bridge -- lookup, not calculation
    Dim gi As Variant, jTxt As String
    jTxt = pick(1)
    If Len(jTxt) = 0 Then FailNow Msg("E_INCOMPLETE")
    gi = Application.Match(jTxt, GridRange("CODE_J_TEXT"), 0)
    If IsError(gi) Then FailNow Msg("E_DATE")

    ' (d) event type = the workbook's own rule tables, in the agreed order:
    '     disposition picked as a terminal state wins, else the station rule
    et = DispRule(code(8))
    If Len(et) = 0 Then et = StationRule(code(3))
    If Len(et) = 0 Then FailNow Msg("E_NORULE")

    ' (e) duplicate barcode -> flagged, never blocked, never deleted
    Dim dups As Long
    dups = Application.WorksheetFunction.CountIf(QcCol(wsD, "BARCODE"), pick(bc))
    status = GridVal(GRID_STAT, ST_OK)
    If dups > 0 Then status = GridVal(GRID_STAT, ST_DUP)

    ' (f) append
    r = NextDataRow(wsD)
    If r = 0 Then FailNow Msg("E_CAPACITY")
    WriteRow wsD, r, jTxt, CLng(GridAt("QC_CAL_G", CLng(gi))), _
             GridAt("QC_CAL_MK", CLng(gi)), GridAt("QC_CAL_WK", CLng(gi)), _
             code(2), code(3), disp(3), code(4), disp(4), code(5), disp(5), _
             code(6), disp(6), code(7), disp(7), code(8), et, pick(bc), _
             MoldCheck(disp(6), pick(bc)), status, Label("CHANNEL_FORM", "FORM-MACRO")

    ' (g) fingerprint, then make the form ready for the next part
    LogLine QcText(wsD, "EVENT_ID", r), "APPEND", "BARCODE", "", pick(bc) & " | " & disp(5)
    wsF.Range("C14").ClearContents
    Application.GoTo wsF.Range("C14"), False

    Dim ok As String
    ok = Msg("OK_APPEND") & " " & QcText(wsD, "EVENT_ID", r) & "  (row " & r & ")"
    If dups > 0 Then ok = ok & vbCrLf & Msg("E_DUP_ACT")
    MsgBox ok, vbInformation, "QC-F-14"
    Exit Sub
Fail:
    FailNow "QC-ERR " & Err.Number & ": " & Err.Description & _
            "  [nothing was written]"
End Sub


'-----------------------------------------------------------------------
' 2. QC_Import_Staging -- the macro-free path, promoted into tblQC
'    (STAGING is intentionally not append-only: an imported row is cleared
'     so it cannot be imported twice. That is the only clearing this module
'     ever does, and only on the scratch sheet -- never on tblQC.)
'-----------------------------------------------------------------------
Public Sub QC_Import_Staging()
    Dim wsD As Worksheet, wsS As Worksheet
    Dim r As Long, n As Long, p As Long, i As Long, tgt As Long, doneRows As String
    Dim pick(1 To 9) As String, code(1 To 9) As String, disp(1 To 9) As String
    Dim et As String, gi As Variant, ready As String

    On Error GoTo Fail
    Set wsS = SheetAt("STAGING")
    Set wsD = SheetAt("DATA")
    ready = ""

    For r = 8 To 207
        If InStr(1, CellTxt(wsS, "K" & r), "Import", vbTextCompare) > 0 Then
            ' B..J = date, shift, station, inspector, part, mold, defect,
            ' disposition, barcode  (columns of the scratch sheet)
            pick(1) = CellTxt(wsS, "B" & r)
            pick(2) = CellTxt(wsS, "C" & r)
            pick(3) = CellTxt(wsS, "D" & r)
            pick(4) = CellTxt(wsS, "E" & r)
            pick(5) = CellTxt(wsS, "F" & r)
            pick(6) = CellTxt(wsS, "G" & r)
            pick(7) = CellTxt(wsS, "H" & r)
            pick(8) = CellTxt(wsS, "I" & r)
            pick(9) = RawText(wsS, "J" & r)
            For i = 2 To 8
                code(i) = PickCode(pick(i))
                disp(i) = PickName(pick(i))
                If Len(code(i)) = 0 Then FailNow Msg("E_INCOMPLETE") & " (row " & r & ")"
            Next i
            If Len(pick(9)) <> BC_LEN Or Not AllDigits(pick(9)) Then _
                FailNow Msg("E_BARCODE") & " (row " & r & ")"
            gi = Application.Match(pick(1), GridRange("CODE_J_TEXT"), 0)
            If IsError(gi) Then FailNow Msg("E_DATE") & " (row " & r & ")"
            et = DispRule(code(8))
            If Len(et) = 0 Then et = StationRule(code(3))
            If Len(et) = 0 Then FailNow Msg("E_NORULE") & " (row " & r & ")"
            p = Application.WorksheetFunction.CountIf(QcCol(wsD, "BARCODE"), pick(9))
            tgt = NextDataRow(wsD)
            If tgt = 0 Then FailNow Msg("E_CAPACITY")
            WriteRow wsD, tgt, pick(1), CLng(GridAt("QC_CAL_G", CLng(gi))), _
                     GridAt("QC_CAL_MK", CLng(gi)), GridAt("QC_CAL_WK", CLng(gi)), _
                     code(2), code(3), disp(3), code(4), disp(4), code(5), disp(5), _
                     code(6), disp(6), code(7), disp(7), code(8), et, pick(9), _
                     MoldCheck(disp(6), pick(9)), _
                     GridVal(GRID_STAT, IIf(p > 0, ST_DUP, ST_OK)), _
                     Label("CHANNEL_STAGING", "STAGING-MACRO")
            wsS.Range("B" & r & ":L" & r).ClearContents
            n = n + 1
            doneRows = doneRows & IIf(Len(doneRows) = 0, "", ", ") & tgt
        End If
    Next r

    If n = 0 Then
        MsgBox Msg("E_STAGE_NONE"), vbExclamation, "QC-F-14"
        Exit Sub
    End If
    LogLine "STAGING", "IMPORT", "ROWS", "", n & " rows -> " & doneRows
    MsgBox Msg("OK_IMPORT") & "  (" & n & ")", vbInformation, "QC-F-14"
    Exit Sub
Fail:
    FailNow "QC-ERR " & Err.Number & ": " & Err.Description & " [import stopped]"
End Sub


'-----------------------------------------------------------------------
' 3. QC_VoidRow -- a mistake is voided, never erased
'-----------------------------------------------------------------------
Public Sub QC_VoidRow()
    Dim wsD As Worksheet, lo As ListObject, r As Long
    On Error GoTo Fail
    Set wsD = SheetAt("DATA")
    Set lo = wsD.ListObjects(TBL_QC)
    If Selection.Rows.Count <> 1 Or Selection.Areas.Count <> 1 Or _
       Intersect(Selection, lo.DataBodyRange) Is Nothing Then FailNow Msg("E_VOID")
    r = Selection.Row
    If QcText(wsD, "ROW_STATUS", r) = GridVal(GRID_STAT, ST_VOID) Then
        MsgBox Msg("E_VOID"), vbExclamation, "QC-F-14"
        Exit Sub
    End If
    PutText wsD, r, "ROW_STATUS", GridVal(GRID_STAT, ST_VOID)
    LogLine QcText(wsD, "EVENT_ID", r), "VOID", "ROW_STATUS", _
            QcText(wsD, "ROW_STATUS", r), GridVal(GRID_STAT, ST_VOID)
    MsgBox Msg("OK_VOID"), vbInformation, "QC-F-14"
    Exit Sub
Fail:
    FailNow "QC-ERR " & Err.Number & ": " & Err.Description
End Sub


'-----------------------------------------------------------------------
' 4. Small, boring utilities
'-----------------------------------------------------------------------
Public Sub QC_ClearForm()
    ' only the nine input cells; the label cells in columns B and E stay put
    Dim wsF As Worksheet
    On Error Resume Next
    Set wsF = SheetAt("FORM")
    If wsF Is Nothing Then Exit Sub
    Union(wsF.Range("C6,C8,C10,C12,C14"), wsF.Range("F6,F8,F10,F12")).ClearContents
    wsF.Range("C14").Select
End Sub

Public Sub QC_About()
    Dim n As Long
    On Error Resume Next
    n = Application.WorksheetFunction.CountA(SheetAt("DATA").ListObjects(TBL_QC).ListColumns(1).DataBodyRange)
    MsgBox "QC-F-14 " & QC_VERSION & vbCrLf & Msg("ABOUT") & vbCrLf & vbCrLf & _
           "tblQC rows: " & n & "  |  user: " & UserName() & "  |  box: " & Environ$("COMPUTERNAME"), _
           vbInformation, "QC-F-14"
End Sub

' -- writing one complete fact row (24 value columns + 2 audit columns) ----
Private Sub WriteRow(ByVal wsD As Worksheet, ByVal r As Long, ByVal jTxt As String, _
                     ByVal gSer As Long, ByVal mKey As Variant, ByVal wKey As Variant, _
                     ByVal shiftId As String, ByVal stCode As String, ByVal stName As String, _
                     ByVal inspId As String, ByVal inspName As String, _
                     ByVal partId As String, ByVal partName As String, _
                     ByVal moldId As String, ByVal moldName As String, _
                     ByVal defId As String, ByVal defName As String, _
                     ByVal dispId As String, ByVal et As String, ByVal bc As String, _
                     ByVal mc As String, ByVal status As String, ByVal channel As String)
    Dim n As Long
    n = Application.WorksheetFunction.CountA(QcCol(wsD, "EVENT_ID")) + 1
    PutText wsD, r, "EVENT_ID", "E-" & Format$(n, "000000")
    PutText wsD, r, "EVENT_DATE_J", jTxt
    PutNum wsD, r, "EVENT_DATE_G", gSer, "yyyy/mm/dd"
    PutNum wsD, r, "J_MONTH_KEY", CLng(mKey), "0"
    PutText wsD, r, "J_WEEK_KEY", CStr(wKey)
    PutText wsD, r, "SHIFT_ID", shiftId
    PutText wsD, r, "STATION_ID", stCode
    PutText wsD, r, "STATION_SNAP", stName
    PutText wsD, r, "PART_ID", partId
    PutText wsD, r, "PART_NAME_SNAP", partName
    PutText wsD, r, "MOLD_ID", moldId
    PutText wsD, r, "MOLD_NAME_SNAP", moldName
    PutText wsD, r, "DEFECT_ID", defId
    PutText wsD, r, "DEFECT_NAME_SNAP", defName
    PutText wsD, r, "DISPOSITION_ID", dispId
    PutText wsD, r, "EVENT_TYPE", et
    PutText wsD, r, "BARCODE", bc                    ' Text format first -> 0-preserved
    PutText wsD, r, "BC_PART_CODE", Left$(bc, 2)
    PutText wsD, r, "MOLD_CHECK", mc
    PutNum wsD, r, "ENTRY_STAMP", Now, "yyyy/mm/dd hh:mm"
    PutText wsD, r, "ENTRY_USER", UserName()        ' who pressed the button (OS account)
    PutText wsD, r, "INSPECTOR_ID", inspId           ' who inspected  (form field 4 / Master)
    PutText wsD, r, "INSPECTOR_SNAP", inspName
    PutText wsD, r, "ENTRY_CHANNEL", channel
    PutText wsD, r, "ROW_STATUS", status
    PutText wsD, r, "NOTE", ""                       ' never leave stale text in a reused row
End Sub

Private Sub PutText(ByVal ws As Worksheet, ByVal r As Long, ByVal token As String, _
                    ByVal v As String)
    Dim c As Range
    Set c = ws.Cells(r, ColIdx(ws, TBL_QC, token))
    c.NumberFormat = "@"
    c.Value = v
End Sub

Private Sub PutNum(ByVal ws As Worksheet, ByVal r As Long, ByVal token As String, _
                   ByVal v As Variant, ByVal fmt As String)
    With ws.Cells(r, ColIdx(ws, TBL_QC, token))
        .NumberFormat = fmt
        .Value = v
    End With
End Sub

' -- reading: header token -> column index, so no column letter is hardcoded
Private Function ColIdx(ByVal ws As Worksheet, ByVal tbl As String, ByVal token As String) As Long
    Dim lo As ListObject, i As Long
    Set lo = ws.ListObjects(tbl)
    For i = 1 To lo.ListColumns.Count
        If InStr(1, lo.ListColumns(i).Name, token, vbTextCompare) > 0 Then
            ColIdx = lo.ListColumns(i).Index
            Exit Function
        End If
    Next i
    Err.Raise vbObjectError + 520, "QC", "column '" & token & "' not found in " & tbl
End Function

Private Function QcCol(ByVal ws As Worksheet, ByVal token As String) As Range
    Set QcCol = ws.ListObjects(TBL_QC).ListColumns(ColIdx(ws, TBL_QC, token)).DataBodyRange
End Function

Private Function QcText(ByVal ws As Worksheet, ByVal token As String, ByVal r As Long) As String
    QcText = Trim$(CStr(ws.Cells(r, ColIdx(ws, TBL_QC, token)).Value))
End Function

Private Function NextDataRow(ByVal ws As Worksheet) As Long
    Dim lo As ListObject, r As Long
    Set lo = ws.ListObjects(TBL_QC)
    r = lo.HeaderRowRange.Row + 1
    Do While r <= lo.Range.Row + lo.Range.Rows.Count - 1
        If Len(Trim$(CStr(ws.Cells(r, 1).Value))) = 0 Then
            NextDataRow = r
            Exit Function
        End If
        r = r + 1
    Loop
    NextDataRow = 0                 ' table is full -- caller must stop, not extend
End Function

Private Function CellTxt(ByVal ws As Worksheet, ByVal addr As String) As String
    CellTxt = Trim$(CStr(ws.Range(addr).Value))
End Function

' The barcode is read with .Text so a lost leading zero shows up as a short
' string (and is then rejected) instead of silently becoming a 18-digit value.
Private Function RawText(ByVal ws As Worksheet, ByVal addr As String) As String
    RawText = Trim$(ws.Range(addr).Text)
End Function

Private Function GridRange(ByVal nm As String) As Range
    On Error Resume Next
    Set GridRange = ThisWorkbook.Names(nm).RefersToRange
    On Error GoTo 0
End Function

Private Function GridAt(ByVal nm As String, ByVal i As Long) As Variant
    GridAt = GridRange(nm).Cells(i, 1).Value
End Function

Private Function GridVal(ByVal nm As String, ByVal i As Long) As String
    GridVal = Trim$(CStr(GridRange(nm).Cells(i, 1).Value))
End Function

' -- Settings!QC_MSG / QC_SHEETS: the macro has no hardcoded name or message
Private Function Cfg(ByVal nm As String, ByVal k As String) As String
    Dim rg As Range, i As Variant
    Set rg = GridRange(nm)
    If rg Is Nothing Then Err.Raise vbObjectError + 521, "QC", "name '" & nm & "' missing"
    i = Application.Match(k, rg.Columns(1), 0)
    If IsError(i) Then
        Cfg = ""
    Else
        Cfg = Trim$(CStr(rg.Cells(CLng(i), 2).Value))
    End If
End Function

Private Function Msg(ByVal k As String) As String
    Dim s As String
    s = Cfg(CFG_MSG, k)
    If Len(s) = 0 Then s = k          ' never show an empty box if the row is missing
    Msg = s
End Function

Private Function Label(ByVal k As String, ByVal fallback As String) As String
    Dim s As String
    s = Cfg(CFG_MSG, k)
    If Len(s) = 0 Then s = fallback
    Label = s
End Function

Private Function SheetAt(ByVal k As String) As Worksheet
    Dim s As String
    s = Cfg(CFG_SHEETS, k)
    If Len(s) = 0 Then Err.Raise vbObjectError + 522, "QC", Msg("E_CONFIG") & " (" & k & ")"
    On Error Resume Next
    Set SheetAt = ThisWorkbook.Worksheets(s)
    On Error GoTo 0
    If SheetAt Is Nothing Then Err.Raise vbObjectError + 523, "QC", Msg("E_SHEET") & " (" & s & ")"
End Function

' "P-29 | 470 fender" -> "P-29"   /   -> "470 fender"
Private Function PickCode(ByVal s As String) As String
    Dim p As Long
    s = Trim$(s)
    p = InStr(s, "|")
    If p = 0 Then
        PickCode = s
    Else
        PickCode = Trim$(Left$(s, p - 1))
    End If
End Function

Private Function PickName(ByVal s As String) As String
    Dim p As Long
    s = Trim$(s)
    p = InStr(s, "|")
    If p > 0 Then PickName = Trim$(Mid$(s, p + 1))
End Function

' -- the two rule tables, read from Master (never duplicated in code) -----
Private Function DispRule(ByVal dispCode As String) As String
    Dim ws As Worksheet, lo As ListObject, i As Long, v As Variant
    Set ws = SheetAt("MASTER")
    Set lo = ws.ListObjects(TBL_DISP)
    For i = 1 To lo.ListRows.Count
        If Trim$(CStr(lo.ListColumns(1).DataBodyRange.Cells(i, 1).Value)) = dispCode Then
            v = lo.ListColumns(lo.ListColumns.Count).DataBodyRange.Cells(i, 1).Value
            DispRule = Trim$(CStr(v))     ' empty = "no terminal event type for this pick"
            Exit Function
        End If
    Next i
End Function

Private Function StationRule(ByVal stCode As String) As String
    Dim ws As Worksheet, lo As ListObject, i As Long
    Set ws = SheetAt("MASTER")
    Set lo = ws.ListObjects(TBL_RULE)
    For i = 1 To lo.ListRows.Count
        If Trim$(CStr(lo.ListColumns(1).DataBodyRange.Cells(i, 1).Value)) = stCode Then
            StationRule = Trim$(CStr(lo.ListColumns(3).DataBodyRange.Cells(i, 1).Value))
            Exit Function
        End If
    Next i
End Function

' cavity vs barcode digit 4 -- mirrors tools/xlmigrate.barcode_semantics();
' a warning column, never a reason to refuse a record.
Private Function MoldCheck(ByVal moldName As String, ByVal bc As String) As String
    Dim mv As String, g As Long
    mv = UCase$(Trim$(moldName))
    If Len(mv) <> 1 Then
        MoldCheck = GridVal(GRID_MOLD, MC_DIRECTIONAL)
    ElseIf Asc(mv) < 65 Or Asc(mv) > 69 Then
        MoldCheck = GridVal(GRID_MOLD, MC_DIRECTIONAL)
    Else
        g = Val(Mid$(bc, 4, 1))
        If g < 1 Or g > 5 Then
            MoldCheck = GridVal(GRID_MOLD, MC_DIGIT)
        ElseIf mv = Chr$(64 + g) Then
            MoldCheck = GridVal(GRID_MOLD, MC_MATCH)
        Else
            MoldCheck = GridVal(GRID_MOLD, MC_MISS)
        End If
    End If
End Function

Private Function AllDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Asc(Mid$(s, i, 1)) < 48 Or Asc(Mid$(s, i, 1)) > 57 Then Exit Function
    Next i
    AllDigits = True
End Function

' -- best-effort fingerprint; a full ring never blocks a registration -----
Private Sub LogLine(ByVal eid As String, ByVal action As String, ByVal field As String, _
                    ByVal oldV As String, ByVal newV As String)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = SheetAt("LOG")
    If ws Is Nothing Then Exit Sub
    r = LOG_FIRST
    Do While r <= LOG_LAST
        If Len(Trim$(CStr(ws.Cells(r, 1).Value))) = 0 Then Exit Do
        r = r + 1
    Loop
    If r > LOG_LAST Then Exit Sub          ' ring full: keep the data, lose the note
    ws.Cells(r, 1).NumberFormat = "@"
    ws.Cells(r, 1).Value = eid
    ws.Cells(r, 2).Value = action
    ws.Cells(r, 3).Value = UserName()
    ws.Cells(r, 4).Value = Environ$("COMPUTERNAME")
    ws.Cells(r, 5).NumberFormat = "yyyy/mm/dd hh:mm"
    ws.Cells(r, 5).Value = Now
    ws.Cells(r, 6).Value = field
    ws.Cells(r, 7).NumberFormat = "@"
    ws.Cells(r, 7).Value = oldV
    ws.Cells(r, 8).NumberFormat = "@"
    ws.Cells(r, 8).Value = newV
End Sub

Private Sub FailNow(ByVal s As String)
    Beep
    MsgBox s, vbExclamation, "QC-F-14 -- nothing was written"
    Err.Clear
End Sub
