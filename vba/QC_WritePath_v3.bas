Attribute VB_Name = "QC_WritePath_v3"
Option Explicit
'=======================================================================
'  QC-F-14  v3.0  --  WRITE PATH ONLY  (v3: Record_ID QC-..., Record_Type,
'  Inspection_Result, Disposition 5 values, Current_Status 5 values)
'-----------------------------------------------------------------------
'  ARCHITECTURE:
'    * formulas compute, VBA writes. No KPI, no pivot, no delete.
'    * Record_ID = QC- + 8 digits, auto MAX+1, immutable.
'    * Barcode is Text 19, leading zeros preserved. Duplicate NOT blocked:
'      flagged as REINSPECTION + CURRENT_STATUS=OPEN / ROW_STATUS=suspect.
'    * Record_Type: RETURN if station = برگشت از فروش, else REINSPECTION if
'      barcode exists, else INSPECTION. Determined from Data, not guessed.
'    * Inspection_Result: OK/NOK from form. RULE_LOGIC checks OK+NOK vs defect.
'    * Disposition: ACCEPT/REWORK/SCRAP/RETURN_TO_PROCESS/HOLD (5 values)
'    * Current_Status: OPEN/UNDER_REWORK/WAITING_REINSPECTION/CLOSED/VOID
'    * Void: QC_VoidRow sets CURRENT_STATUS=VOID, VOID_FLAG=TRUE, asks reason.
'    * All Persian strings read from Settings!QC_MSG and Helper grids (ASCII only).
'=======================================================================

Private Const TBL_QC As String = "tblQC"
Private Const TBL_RULE As String = "RULE_STATION"
Private Const TBL_DISP_OLD As String = "DIM_DISPOSITION"
Private Const GRID_MOLD As String = "CODE_MOLD_CHECK"
Private Const GRID_STAT As String = "CODE_ROW_STATUS"
Private Const GRID_RECTYPE As String = "CODE_RECORDTYPE"
Private Const GRID_RESULT As String = "CODE_RESULT"
Private Const GRID_DISP As String = "CODE_DISPOSITION_NEW"
Private Const GRID_STATUS As String = "CODE_STATUS_NEW"
Private Const CFG_SHEETS As String = "QC_SHEETS"
Private Const CFG_MSG As String = "QC_MSG"
Private Const BC_LEN As Long = 19
Private Const LOG_FIRST As Long = 3
Private Const LOG_LAST As Long = 202

Private Const MC_MATCH As Long = 1
Private Const MC_MISS As Long = 2
Private Const MC_DIRECTIONAL As Long = 3
Private Const MC_DIGIT As Long = 4

Private Const ST_OK As Long = 1
Private Const ST_VOID As Long = 3
Private Const ST_DUP As Long = 4

Private Const RT_INSPECTION As Long = 1
Private Const RT_REINSPECTION As Long = 2
Private Const RT_RETURN As Long = 3

Private Const RES_OK As Long = 1
Private Const RES_NOK As Long = 2

Private Const DISP_ACCEPT As Long = 1
Private Const DISP_REWORK As Long = 2
Private Const DISP_SCRAP As Long = 3
Private Const DISP_RETURN As Long = 4
Private Const DISP_HOLD As Long = 5

Private Const STAT_OPEN As Long = 1
Private Const STAT_UNDER As Long = 2
Private Const STAT_WAIT As Long = 3
Private Const STAT_CLOSED As Long = 4
Private Const STAT_VOID As Long = 5

Public Const QC_VERSION As String = "3.0.0-writepath"

'-----------------------------------------------------------------------
' 1. QC_Append v3
'-----------------------------------------------------------------------
Public Sub QC_Append()
    Dim wsD As Worksheet, wsF As Worksheet
    Dim pickDate As String, pickShift As String, pickStation As String, pickStationName As String
    Dim pickInspector As String, pickInspectorName As String
    Dim pickPart As String, pickPartName As String
    Dim pickMold As String, pickMoldName As String
    Dim pickBarcode As String
    Dim pickResult As String, pickDefect As String, pickDefectName As String
    Dim pickDisp As String, pickStatus As String, pickNote As String
    Dim codeShift As String, codeStation As String, codeInspector As String
    Dim codePart As String, codeMold As String, codeDefect As String
    Dim etLegacy As String, recType As String, curStatus As String
    Dim dups As Long, r As Long

    On Error GoTo Fail
    Set wsF = SheetAt("FORM")
    Set wsD = SheetAt("DATA")

    ' --- read form (v3 layout) ---
    pickDate = CellTxt(wsF, "C6")
    pickShift = CellTxt(wsF, "C8")
    pickStation = CellTxt(wsF, "C10")
    pickInspector = CellTxt(wsF, "C12")
    pickPart = CellTxt(wsF, "F6")
    pickMold = CellTxt(wsF, "F8")
    pickBarcode = RawText(wsF, "C14")
    pickResult = CellTxt(wsF, "F10")
    pickDefect = CellTxt(wsF, "F12")
    pickDisp = CellTxt(wsF, "F14")
    pickStatus = CellTxt(wsF, "C16")
    pickNote = CellTxt(wsF, "F16")

    codeShift = PickCode(pickShift)
    codeStation = PickCode(pickStation)
    pickStationName = PickName(pickStation)
    codeInspector = PickCode(pickInspector)
    pickInspectorName = PickName(pickInspector)
    codePart = PickCode(pickPart)
    pickPartName = PickName(pickPart)
    codeMold = PickCode(pickMold)
    pickMoldName = PickName(pickMold)
    codeDefect = PickCode(pickDefect)
    pickDefectName = PickName(pickDefect)

    ' required: date, shift, station, inspector, part, mold, barcode, result, disposition
    If Len(pickDate) = 0 Or Len(codeShift) = 0 Or Len(codeStation) = 0 Or Len(codeInspector) = 0 _
       Or Len(codePart) = 0 Or Len(codeMold) = 0 Or Len(pickBarcode) = 0 _
       Or Len(PickCode(pickResult)) = 0 Or Len(PickCode(pickDisp)) = 0 Then
        FailNow Msg("E_INCOMPLETE") & vbCrLf & CellTxt(wsF, "C28")
        Exit Sub
    End If

    ' barcode 19 digits
    If Len(pickBarcode) <> BC_LEN Or Not AllDigits(pickBarcode) Then FailNow Msg("E_BARCODE"): Exit Sub

    ' calendar bridge
    Dim gi As Variant
    gi = Application.Match(pickDate, GridRange("CODE_J_TEXT"), 0)
    If IsError(gi) Then FailNow Msg("E_DATE"): Exit Sub

    ' duplicate check for Record_Type
    dups = Application.WorksheetFunction.CountIf(QcCol(wsD, "BARCODE"), pickBarcode)

    ' Record_Type logic: RETURN if station contains برگشت, else REINSPECTION if exists, else INSPECTION
    If InStr(1, pickStationName, "برگشت", vbTextCompare) > 0 Or codeStation = "ST-08" Then
        recType = GridVal(GRID_RECTYPE, RT_RETURN)
    ElseIf dups > 0 Then
        recType = GridVal(GRID_RECTYPE, RT_REINSPECTION)
    Else
        recType = GridVal(GRID_RECTYPE, RT_INSPECTION)
    End If

    ' Inspection_Result logic: OK/NOK from form, but enforce RULE_LOGIC
    Dim resCode As String, resName As String
    resCode = PickCode(pickResult)
    resName = PickName(pickResult)
    If Len(resName) = 0 Then resName = resCode
    ' If OK but defect present -> block
    If UCase$(resName) = "OK" And Len(codeDefect) > 0 Then
        FailNow "RULE-01: OK with defect not allowed"
        Exit Sub
    End If
    If UCase$(resName) = "NOK" And Len(codeDefect) = 0 Then
        FailNow "RULE-02: NOK without defect not allowed (use HOLD if pending)"
        Exit Sub
    End If

    ' Disposition: from form, already code
    Dim dispCode As String, dispName As String
    dispCode = PickCode(pickDisp)
    dispName = PickName(pickDisp)
    If Len(dispName) = 0 Then dispName = dispCode

    ' Current_Status: from form if provided, else derive
    If Len(PickCode(pickStatus)) > 0 Then
        curStatus = PickName(pickStatus)
        If Len(curStatus) = 0 Then curStatus = PickCode(pickStatus)
    Else
        ' derive: REWORK->UNDER_REWORK, HOLD->OPEN, else CLOSED
        If dispName = "REWORK" Then
            curStatus = GridVal(GRID_STATUS, STAT_UNDER)
        ElseIf dispName = "HOLD" Then
            curStatus = GridVal(GRID_STATUS, STAT_OPEN)
        Else
            curStatus = GridVal(GRID_STATUS, STAT_CLOSED)
        End If
    End If

    ' legacy EVENT_TYPE for backward compat
    etLegacy = DispRuleOld(codeStation)
    If Len(etLegacy) = 0 Then etLegacy = StationRule(codeStation)
    If Len(etLegacy) = 0 Then etLegacy = "اصلاحی"

    ' status for duplicate
    Dim rowStat As String
    rowStat = GridVal(GRID_STAT, ST_OK)
    If dups > 0 Then rowStat = GridVal(GRID_STAT, ST_DUP)

    r = NextDataRow(wsD)
    If r = 0 Then FailNow Msg("E_CAPACITY"): Exit Sub

    WriteRowV3 wsD, r, pickDate, CLng(GridAt("QC_CAL_G", CLng(gi))), _
               GridAt("QC_CAL_MK", CLng(gi)), GridAt("QC_CAL_WK", CLng(gi)), _
               codeShift, codeStation, pickStationName, _
               codeInspector, pickInspectorName, _
               codePart, pickPartName, _
               codeMold, pickMoldName, _
               pickBarcode, MoldCheck(pickMoldName, pickBarcode), _
               resName, codeDefect, pickDefectName, _
               dispName, curStatus, _
               rowStat, pickNote, recType, etLegacy

    LogLine QcText(wsD, "RECORD_ID", r), "APPEND", "BARCODE", "", pickBarcode & " | " & recType & " | " & dispName

    wsF.Range("C14").ClearContents
    wsF.Range("F16").ClearContents
    Application.GoTo wsF.Range("C14"), False

    Dim ok As String
    ok = Msg("OK_APPEND") & " " & QcText(wsD, "RECORD_ID", r) & " (" & recType & ")"
    If dups > 0 Then ok = ok & vbCrLf & Msg("E_DUP_ACT")
    MsgBox ok, vbInformation, "QC-F-14 v3"
    Exit Sub
Fail:
    FailNow "QC-ERR " & Err.Number & ": " & Err.Description & " [nothing written]"
End Sub

'-----------------------------------------------------------------------
' 2. Import Staging v3
'-----------------------------------------------------------------------
Public Sub QC_Import_Staging()
    Dim wsD As Worksheet, wsS As Worksheet
    Dim r As Long, n As Long, tgt As Long, dups As Long
    Dim pickDate As String, codeShift As String, codeStation As String, pickStationName As String
    Dim codeInspector As String, pickInspectorName As String
    Dim codePart As String, pickPartName As String
    Dim codeMold As String, pickMoldName As String
    Dim pickBarcode As String, resName As String, codeDefect As String, pickDefectName As String
    Dim dispName As String, curStatus As String, gi As Variant, recType As String

    On Error GoTo Fail
    Set wsS = SheetAt("STAGING")
    Set wsD = SheetAt("DATA")

    For r = 8 To 207
        If InStr(1, CellTxt(wsS, "L" & r), "Import", vbTextCompare) > 0 Then
            pickDate = CellTxt(wsS, "B" & r)
            codeShift = PickCode(CellTxt(wsS, "C" & r))
            codeStation = PickCode(CellTxt(wsS, "D" & r))
            pickStationName = PickName(CellTxt(wsS, "D" & r))
            codeInspector = PickCode(CellTxt(wsS, "E" & r))
            pickInspectorName = PickName(CellTxt(wsS, "E" & r))
            codePart = PickCode(CellTxt(wsS, "F" & r))
            pickPartName = PickName(CellTxt(wsS, "F" & r))
            codeMold = PickCode(CellTxt(wsS, "G" & r))
            pickMoldName = PickName(CellTxt(wsS, "G" & r))
            pickBarcode = RawText(wsS, "H" & r)
            resName = PickName(CellTxt(wsS, "I" & r))
            If Len(resName) = 0 Then resName = PickCode(CellTxt(wsS, "I" & r))
            codeDefect = PickCode(CellTxt(wsS, "J" & r))
            pickDefectName = PickName(CellTxt(wsS, "J" & r))
            dispName = PickName(CellTxt(wsS, "K" & r))
            If Len(dispName) = 0 Then dispName = PickCode(CellTxt(wsS, "K" & r))
            curStatus = PickName(CellTxt(wsS, "L" & r))

            If Len(pickDate) = 0 Or Len(codeShift) = 0 Or Len(codeStation) = 0 Or Len(pickBarcode) = 0 Then
                FailNow Msg("E_INCOMPLETE") & " row " & r: Exit Sub
            End If
            If Len(pickBarcode) <> BC_LEN Or Not AllDigits(pickBarcode) Then FailNow Msg("E_BARCODE") & " row " & r: Exit Sub
            gi = Application.Match(pickDate, GridRange("CODE_J_TEXT"), 0)
            If IsError(gi) Then FailNow Msg("E_DATE") & " row " & r: Exit Sub

            dups = Application.WorksheetFunction.CountIf(QcCol(wsD, "BARCODE"), pickBarcode)
            If InStr(1, pickStationName, "برگشت", vbTextCompare) > 0 Then
                recType = GridVal(GRID_RECTYPE, RT_RETURN)
            ElseIf dups > 0 Then
                recType = GridVal(GRID_RECTYPE, RT_REINSPECTION)
            Else
                recType = GridVal(GRID_RECTYPE, RT_INSPECTION)
            End If
            If Len(curStatus) = 0 Then
                If dispName = "REWORK" Then curStatus = GridVal(GRID_STATUS, STAT_UNDER) Else curStatus = GridVal(GRID_STATUS, STAT_CLOSED)
            End If

            tgt = NextDataRow(wsD)
            If tgt = 0 Then FailNow Msg("E_CAPACITY"): Exit Sub

            WriteRowV3 wsD, tgt, pickDate, CLng(GridAt("QC_CAL_G", CLng(gi))), _
                       GridAt("QC_CAL_MK", CLng(gi)), GridAt("QC_CAL_WK", CLng(gi)), _
                       codeShift, codeStation, pickStationName, _
                       codeInspector, pickInspectorName, _
                       codePart, pickPartName, _
                       codeMold, pickMoldName, _
                       pickBarcode, MoldCheck(pickMoldName, pickBarcode), _
                       resName, codeDefect, pickDefectName, _
                       dispName, curStatus, _
                       GridVal(GRID_STAT, IIf(dups > 0, ST_DUP, ST_OK)), "", recType, ""

            wsS.Range("B" & r & ":L" & r).ClearContents
            n = n + 1
        End If
    Next r
    If n = 0 Then MsgBox Msg("E_STAGE_NONE"), vbExclamation, "QC-F-14": Exit Sub
    LogLine "STAGING", "IMPORT", "ROWS", "", CStr(n)
    MsgBox Msg("OK_IMPORT") & " (" & n & ")", vbInformation, "QC-F-14"
    Exit Sub
Fail:
    FailNow "QC-ERR " & Err.Number & ": " & Err.Description & " [import stopped]"
End Sub

'-----------------------------------------------------------------------
' 3. Void
'-----------------------------------------------------------------------
Public Sub QC_VoidRow()
    Dim wsD As Worksheet, lo As ListObject, r As Long, reason As String
    On Error GoTo Fail
    Set wsD = SheetAt("DATA")
    Set lo = wsD.ListObjects(TBL_QC)
    If Selection.Rows.Count <> 1 Or Selection.Areas.Count <> 1 Or Intersect(Selection, lo.DataBodyRange) Is Nothing Then
        FailNow Msg("E_VOID"): Exit Sub
    End If
    r = Selection.Row
    reason = InputBox("دلیل ابطال را وارد کنید (الزامی):", "Void Reason")
    If Len(Trim$(reason)) = 0 Then MsgBox "ابطال بدون دلیل مجاز نیست (RULE-07)", vbExclamation: Exit Sub
    PutText wsD, r, "CURRENT_STATUS", GridVal(GRID_STATUS, STAT_VOID)
    PutText wsD, r, "VOID_FLAG", "TRUE"
    PutText wsD, r, "VOID_REASON", reason
    PutText wsD, r, "ROW_STATUS", GridVal(GRID_STAT, ST_VOID)
    LogLine QcText(wsD, "RECORD_ID", r), "VOID", "VOID_REASON", "", reason
    MsgBox Msg("OK_VOID"), vbInformation, "QC-F-14"
    Exit Sub
Fail:
    FailNow "QC-ERR " & Err.Number & ": " & Err.Description
End Sub

Public Sub QC_ClearForm()
    Dim wsF As Worksheet
    On Error Resume Next
    Set wsF = SheetAt("FORM")
    If wsF Is Nothing Then Exit Sub
    Union(wsF.Range("C6,C8,C10,C12,C14,C16"), wsF.Range("F6,F8,F10,F12,F14,F16")).ClearContents
    wsF.Range("C14").Select
End Sub

Public Sub QC_About()
    Dim n As Long
    On Error Resume Next
    n = Application.WorksheetFunction.CountA(SheetAt("DATA").ListObjects(TBL_QC).ListColumns(1).DataBodyRange)
    MsgBox "QC-F-14 " & QC_VERSION & vbCrLf & Msg("ABOUT") & vbCrLf & "rows: " & n, vbInformation, "QC-F-14"
End Sub

'--- write full v3 row ---
Private Sub WriteRowV3(ByVal wsD As Worksheet, ByVal r As Long, ByVal jTxt As String, _
                       ByVal gSer As Long, ByVal mKey As Variant, ByVal wKey As Variant, _
                       ByVal shiftId As String, ByVal stCode As String, ByVal stName As String, _
                       ByVal inspId As String, ByVal inspName As String, _
                       ByVal partId As String, ByVal partName As String, _
                       ByVal moldId As String, ByVal moldName As String, _
                       ByVal bc As String, ByVal mc As String, _
                       ByVal res As String, ByVal defId As String, ByVal defName As String, _
                       ByVal disp As String, ByVal curStat As String, _
                       ByVal rowStat As String, ByVal note As String, ByVal recType As String, ByVal etLegacy As String)
    Dim n As Long
    n = Application.WorksheetFunction.CountA(QcCol(wsD, "RECORD_ID")) + 1
    PutText wsD, r, "RECORD_ID", "QC-" & Format$(n, "00000000")
    PutText wsD, r, "RECORD_TYPE", recType
    PutText wsD, r, "EVENT_DATE_J", jTxt
    PutNum wsD, r, "EVENT_DATE_G", gSer, "yyyy/mm/dd"
    PutNum wsD, r, "J_MONTH_KEY", CLng(mKey), "0"
    PutText wsD, r, "J_WEEK_KEY", CStr(wKey)
    PutText wsD, r, "SHIFT_ID", shiftId
    PutText wsD, r, "STATION_ID", stCode
    PutText wsD, r, "STATION_SNAP", stName
    PutText wsD, r, "INSPECTOR_ID", inspId
    PutText wsD, r, "INSPECTOR_SNAP", inspName
    PutText wsD, r, "PART_ID", partId
    PutText wsD, r, "PART_NAME_SNAP", partName
    PutText wsD, r, "MOLD_ID", moldId
    PutText wsD, r, "MOLD_NAME_SNAP", moldName
    PutText wsD, r, "BARCODE", bc
    PutText wsD, r, "BC_PART_CODE", Left$(bc, 2)
    PutText wsD, r, "MOLD_CHECK", mc
    PutText wsD, r, "INSPECTION_RESULT", res
    PutText wsD, r, "DEFECT_ID", defId
    PutText wsD, r, "DEFECT_NAME_SNAP", defName
    PutText wsD, r, "DISPOSITION", disp
    PutText wsD, r, "CURRENT_STATUS", curStat
    PutText wsD, r, "ENTRY_USER", UserName()
    PutNum wsD, r, "ENTRY_STAMP", Now, "yyyy/mm/dd hh:mm"
    PutText wsD, r, "ENTRY_CHANNEL", Label("CHANNEL_FORM", "FORM-MACRO")
    PutText wsD, r, "ROW_STATUS", rowStat
    PutText wsD, r, "NOTE", note
    PutText wsD, r, "VOID_FLAG", "FALSE"
    PutText wsD, r, "VOID_REASON", ""
    PutText wsD, r, "EVENT_TYPE", etLegacy
    PutText wsD, r, "DISPOSITION_ID", ""
End Sub

Private Sub PutText(ByVal ws As Worksheet, ByVal r As Long, ByVal token As String, ByVal v As String)
    Dim c As Range
    Set c = ws.Cells(r, ColIdx(ws, TBL_QC, token))
    c.NumberFormat = "@"
    c.Value = v
End Sub

Private Sub PutNum(ByVal ws As Worksheet, ByVal r As Long, ByVal token As String, ByVal v As Variant, ByVal fmt As String)
    With ws.Cells(r, ColIdx(ws, TBL_QC, token))
        .NumberFormat = fmt
        .Value = v
    End With
End Sub

Private Function ColIdx(ByVal ws As Worksheet, ByVal tbl As String, ByVal token As String) As Long
    Dim lo As ListObject, i As Long
    Set lo = ws.ListObjects(tbl)
    For i = 1 To lo.ListColumns.Count
        If InStr(1, lo.ListColumns(i).Name, token, vbTextCompare) > 0 Then
            ColIdx = lo.ListColumns(i).Index
            Exit Function
        End If
    Next i
    Err.Raise vbObjectError + 520, "QC", "column '" & token & "' not found"
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
        If Len(Trim$(CStr(ws.Cells(r, 1).Value))) = 0 Then NextDataRow = r: Exit Function
        r = r + 1
    Loop
    NextDataRow = 0
End Function

Private Function CellTxt(ByVal ws As Worksheet, ByVal addr As String) As String
    CellTxt = Trim$(CStr(ws.Range(addr).Value))
End Function

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

Private Function Cfg(ByVal nm As String, ByVal k As String) As String
    Dim rg As Range, i As Variant
    Set rg = GridRange(nm)
    If rg Is Nothing Then Err.Raise vbObjectError + 521, "QC", "name '" & nm & "' missing"
    i = Application.Match(k, rg.Columns(1), 0)
    If IsError(i) Then Cfg = "" Else Cfg = Trim$(CStr(rg.Cells(CLng(i), 2).Value))
End Function

Private Function Msg(ByVal k As String) As String
    Dim s As String
    s = Cfg(CFG_MSG, k)
    If Len(s) = 0 Then s = k
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

Private Function PickCode(ByVal s As String) As String
    Dim p As Long
    s = Trim$(s)
    p = InStr(s, "|")
    If p = 0 Then PickCode = s Else PickCode = Trim$(Left$(s, p - 1))
End Function

Private Function PickName(ByVal s As String) As String
    Dim p As Long
    s = Trim$(s)
    p = InStr(s, "|")
    If p > 0 Then PickName = Trim$(Mid$(s, p + 1))
End Function

Private Function DispRuleOld(ByVal dispCode As String) As String
    Dim ws As Worksheet, lo As ListObject, i As Long
    Set ws = SheetAt("MASTER")
    On Error Resume Next
    Set lo = ws.ListObjects(TBL_DISP_OLD)
    On Error GoTo 0
    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If Trim$(CStr(lo.ListColumns(1).DataBodyRange.Cells(i, 1).Value)) = dispCode Then
            DispRuleOld = Trim$(CStr(lo.ListColumns(lo.ListColumns.Count).DataBodyRange.Cells(i, 1).Value))
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

Private Sub LogLine(ByVal eid As String, ByVal action As String, ByVal field As String, ByVal oldV As String, ByVal newV As String)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = SheetAt("LOG")
    If ws Is Nothing Then Exit Sub
    r = LOG_FIRST
    Do While r <= LOG_LAST
        If Len(Trim$(CStr(ws.Cells(r, 1).Value))) = 0 Then Exit Do
        r = r + 1
    Loop
    If r > LOG_LAST Then Exit Sub
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
    MsgBox s, vbExclamation, "QC-F-14 v3 -- nothing written"
    Err.Clear
End Sub

Private Function UserName() As String
    UserName = Environ$("USERNAME")
    If Len(UserName) = 0 Then UserName = Environ$("USER")
End Function
