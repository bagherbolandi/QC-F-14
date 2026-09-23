Attribute VB_Name = "QC_Full_v4_SIMPLE"
Option Explicit
'=======================================================================
'  QC-F-14  v4.2 SIMPLE - No table lookup for login (hardcoded users)
'  Use this if you get Subscript out of range on TBL_USERS
'  Users: admin/admin123 ADMIN, operator/1234 OPERATOR, viewer/viewer123 VIEWER
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

Public Const QC_VERSION As String = "4.2.0-simple"

Public Sub Auto_Open()
    On Error Resume Next
    HideAllExceptLogin
    On Error GoTo 0
End Sub

Public Sub Workbook_Open_Handler()
    Auto_Open
End Sub

Public Sub HideAllExceptLogin()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        If ws.Name = "LOGIN ورود" Or ws.Name = "LOGIN" Then
            ws.Visible = xlSheetVisible
        Else
            ws.Visible = xlSheetVeryHidden
        End If
        On Error GoTo 0
    Next ws
    On Error Resume Next
    Dim wsL As Worksheet
    Set wsL = ThisWorkbook.Worksheets("LOGIN ورود")
    If wsL Is Nothing Then Set wsL = ThisWorkbook.Worksheets("LOGIN")
    If Not wsL Is Nothing Then
        wsL.Range("C12").Value = "(not logged in)"
        wsL.Range("C13").Value = ""
        wsL.Range("C10").Value = "Please login - admin / admin123"
    End If
    On Error GoTo 0
End Sub

Public Sub QC_Login()
    Dim wsL As Worksheet
    Dim userIn As String, passIn As String
    Dim role As String, dispName As String
    
    On Error GoTo Fail
    Set wsL = ThisWorkbook.Worksheets("LOGIN ورود")
    If wsL Is Nothing Then Set wsL = ThisWorkbook.Worksheets("LOGIN")
    If wsL Is Nothing Then
        MsgBox "LOGIN sheet not found", vbExclamation
        Exit Sub
    End If
    
    userIn = Trim$(CStr(wsL.Range("C6").Value))
    passIn = Trim$(CStr(wsL.Range("C8").Value))
    
    If Len(userIn) = 0 Or Len(passIn) = 0 Then
        MsgBox "Username or password empty", vbExclamation
        Exit Sub
    End If
    
    ' HARDCODED USERS - no table access, so no Subscript error
    role = ""
    dispName = userIn
    
    If LCase$(userIn) = "admin" And passIn = "admin123" Then
        role = "ADMIN"
        dispName = "Admin"
    ElseIf LCase$(userIn) = "operator" And passIn = "1234" Then
        role = "OPERATOR"
        dispName = "Operator"
    ElseIf LCase$(userIn) = "viewer" And passIn = "viewer123" Then
        role = "VIEWER"
        dispName = "Viewer"
    ElseIf LCase$(userIn) = "i-01" And passIn = "1234" Then
        role = "OPERATOR"
        dispName = "Inspector 01"
    ElseIf LCase$(userIn) = "i-02" And passIn = "1234" Then
        role = "OPERATOR"
        dispName = "Inspector 02"
    ElseIf LCase$(userIn) = "i-03" And passIn = "1234" Then
        role = "OPERATOR"
        dispName = "Inspector 03"
    ElseIf LCase$(userIn) = "i-04" And passIn = "1234" Then
        role = "OPERATOR"
        dispName = "Inspector 04"
    ElseIf LCase$(userIn) = "i-05" And passIn = "1234" Then
        role = "OPERATOR"
        dispName = "Inspector 05"
    End If
    
    If Len(role) = 0 Then
        wsL.Range("C10").Value = "Invalid username or password"
        MsgBox "Invalid username or password: " & userIn & vbCrLf & "Try admin / admin123", vbExclamation, "Login"
        Exit Sub
    End If
    
    wsL.Range("C12").Value = userIn & " (" & dispName & ")"
    wsL.Range("C13").Value = role
    wsL.Range("C10").Value = "Login OK - " & role
    
    ShowSheetsByRole role
    MsgBox "Login OK - " & role & vbCrLf & userIn, vbInformation, "QC-F-14"
    Exit Sub
Fail:
    MsgBox "Login error " & Err.Number & ": " & Err.Description, vbExclamation
End Sub

Public Sub QC_Logout()
    Dim wsL As Worksheet
    Set wsL = ThisWorkbook.Worksheets("LOGIN ورود")
    If wsL Is Nothing Then Set wsL = ThisWorkbook.Worksheets("LOGIN")
    HideAllExceptLogin
    MsgBox "Logged out - only LOGIN visible", vbInformation
End Sub

Private Sub ShowSheetsByRole(ByVal role As String)
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Select Case UCase$(role)
            Case "ADMIN"
                ws.Visible = xlSheetVisible
            Case "OPERATOR"
                If ws.Name = "LOGIN ورود" Or ws.Name = "LOGIN" Or ws.Name = "FORM فرم ثبت" Or ws.Name = "FORM" Or ws.Name = "STAGING ورود اضطراری" Or ws.Name = "STAGING" Or ws.Name = "Help راهنما" Or ws.Name = "Help" Then
                    ws.Visible = xlSheetVisible
                Else
                    ws.Visible = xlSheetVeryHidden
                End If
            Case "VIEWER"
                If ws.Name = "LOGIN ورود" Or ws.Name = "LOGIN" Or ws.Name = "Dashboard داشبورد" Or ws.Name = "Dashboard" Or ws.Name = "Reports گزارش‌ها" Or ws.Name = "Reports" Or ws.Name = "DQ کیفیت داده" Or ws.Name = "DQ" Or ws.Name = "Help راهنما" Or ws.Name = "Help" Then
                    ws.Visible = xlSheetVisible
                Else
                    ws.Visible = xlSheetVeryHidden
                End If
        End Select
        On Error GoTo 0
    Next ws
    On Error Resume Next
    ThisWorkbook.Worksheets("LOGIN ورود").Activate
    On Error GoTo 0
End Sub

' === WRITE PATH (same) ===
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
    Set wsF = GetSheet("FORM")
    Set wsD = GetSheet("DATA")
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
    If Len(pickDate) = 0 Or Len(codeShift) = 0 Or Len(codeStation) = 0 Or Len(codeInspector) = 0 _
       Or Len(codePart) = 0 Or Len(codeMold) = 0 Or Len(pickBarcode) = 0 _
       Or Len(PickCode(pickResult)) = 0 Or Len(PickCode(pickDisp)) = 0 Then
        MsgBox "Incomplete - " & CellTxt(wsF, "C28"), vbExclamation
        Exit Sub
    End If
    If Len(pickBarcode) <> BC_LEN Or Not AllDigits(pickBarcode) Then MsgBox "Barcode must be 19 digits", vbExclamation: Exit Sub
    Dim gi As Variant
    gi = Application.Match(pickDate, RangeByName("CODE_J_TEXT"), 0)
    If IsError(gi) Then MsgBox "Invalid date", vbExclamation: Exit Sub
    dups = Application.WorksheetFunction.CountIf(QcCol(wsD, "BARCODE"), pickBarcode)
    If InStr(1, pickStationName, "Return", vbTextCompare) > 0 Or codeStation = "ST-08" Then
        recType = GridVal(GRID_RECTYPE, 3)
    ElseIf dups > 0 Then
        recType = GridVal(GRID_RECTYPE, 2)
    Else
        recType = GridVal(GRID_RECTYPE, 1)
    End If
    Dim resName As String
    resName = PickName(pickResult)
    If Len(resName) = 0 Then resName = PickCode(pickResult)
    Dim dispName As String
    dispName = PickName(pickDisp)
    If Len(dispName) = 0 Then dispName = PickCode(pickDisp)
    If Len(PickCode(pickStatus)) > 0 Then
        curStatus = PickName(pickStatus)
        If Len(curStatus) = 0 Then curStatus = PickCode(pickStatus)
    Else
        If dispName = "REWORK" Then curStatus = GridVal(GRID_STATUS, 2) Else curStatus = GridVal(GRID_STATUS, 4)
    End If
    etLegacy = "اصلاحی"
    Dim rowStat As String
    rowStat = GridVal(GRID_STAT, 1)
    If dups > 0 Then rowStat = GridVal(GRID_STAT, 4)
    r = NextDataRow(wsD)
    If r = 0 Then MsgBox "Capacity full", vbExclamation: Exit Sub
    WriteRow wsD, r, pickDate, CLng(GridAt("QC_CAL_G", CLng(gi))), GridAt("QC_CAL_MK", CLng(gi)), GridAt("QC_CAL_WK", CLng(gi)), codeShift, codeStation, pickStationName, codeInspector, pickInspectorName, codePart, pickPartName, codeMold, pickMoldName, pickBarcode, "MATCH", resName, codeDefect, pickDefectName, dispName, curStatus, rowStat, pickNote, recType
    wsF.Range("C14").ClearContents
    MsgBox "Appended " & wsD.Cells(r, 1).Value & " (" & recType & ")", vbInformation
    Exit Sub
Fail:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub WriteRow(ByVal wsD As Worksheet, ByVal r As Long, ByVal jTxt As String, ByVal gSer As Long, ByVal mKey As Variant, ByVal wKey As Variant, ByVal shiftId As String, ByVal stCode As String, ByVal stName As String, ByVal inspId As String, ByVal inspName As String, ByVal partId As String, ByVal partName As String, ByVal moldId As String, ByVal moldName As String, ByVal bc As String, ByVal mc As String, ByVal res As String, ByVal defId As String, ByVal defName As String, ByVal disp As String, ByVal curStat As String, ByVal rowStat As String, ByVal note As String, ByVal recType As String)
    Dim n As Long
    n = Application.WorksheetFunction.CountA(QcCol(wsD, "RECORD_ID")) + 1
    PutText wsD, r, "RECORD_ID", "QC-" & Format$(n, "00000000")
    PutText wsD, r, "RECORD_TYPE", recType
    PutText wsD, r, "EVENT_DATE_J", jTxt
    PutNum wsD, r, "EVENT_DATE_G", gSer, "yyyy/mm/dd"
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
    PutText wsD, r, "INSPECTION_RESULT", res
    PutText wsD, r, "DEFECT_ID", defId
    PutText wsD, r, "DEFECT_NAME_SNAP", defName
    PutText wsD, r, "DISPOSITION", disp
    PutText wsD, r, "CURRENT_STATUS", curStat
    PutText wsD, r, "ENTRY_USER", Environ$("USERNAME")
    PutNum wsD, r, "ENTRY_STAMP", Now, "yyyy/mm/dd hh:mm"
    PutText wsD, r, "ROW_STATUS", rowStat
    PutText wsD, r, "NOTE", note
    PutText wsD, r, "VOID_FLAG", "FALSE"
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
End Function

Private Function QcCol(ByVal ws As Worksheet, ByVal token As String) As Range
    Set QcCol = ws.ListObjects(TBL_QC).ListColumns(ColIdx(ws, TBL_QC, token)).DataBodyRange
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

Private Function RangeByName(ByVal nm As String) As Range
    On Error Resume Next
    Set RangeByName = ThisWorkbook.Names(nm).RefersToRange
    On Error GoTo 0
End Function

Private Function GridAt(ByVal nm As String, ByVal i As Long) As Variant
    GridAt = RangeByName(nm).Cells(i, 1).Value
End Function

Private Function GridVal(ByVal nm As String, ByVal i As Long) As String
    GridVal = Trim$(CStr(RangeByName(nm).Cells(i, 1).Value))
End Function

Private Function GetSheet(ByVal key As String) As Worksheet
    On Error Resume Next
    If UCase$(key) = "FORM" Then
        Set GetSheet = ThisWorkbook.Worksheets("FORM فرم ثبت")
        If GetSheet Is Nothing Then Set GetSheet = ThisWorkbook.Worksheets("FORM")
    ElseIf UCase$(key) = "DATA" Then
        Set GetSheet = ThisWorkbook.Worksheets("Data داده")
        If GetSheet Is Nothing Then Set GetSheet = ThisWorkbook.Worksheets("Data")
    End If
    On Error GoTo 0
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

Private Function AllDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Asc(Mid$(s, i, 1)) < 48 Or Asc(Mid$(s, i, 1)) > 57 Then Exit Function
    Next i
    AllDigits = True
End Function
