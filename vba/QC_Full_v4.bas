Attribute VB_Name = "QC_Full_v4"
Option Explicit
'=======================================================================
'  QC-F-14  v4.0  --  WRITE PATH + ROLE-BASED LOGIN (VeryHidden)
'-----------------------------------------------------------------------
'  SECURITY NOTE: Excel sheet protection / VeryHidden is NOT real
'  security. Passwords are Plain Text in USERS sheet. This is only
'  a deterrent for operators. For real security -> SQL/Web.
'
'  ROLES:
'    ADMIN    -> all sheets visible (Data, Master, Reports, Dashboard,
'                DQ, Settings, Help, LOG_AUDIT, USERS, FORM, STAGING)
'    OPERATOR -> FORM + STAGING + Help + LOGIN only
'    VIEWER   -> Dashboard + Reports + DQ + Help + LOGIN only
'
'  LOGIN SHEET:  LOGIN ورود
'    C6 = username input
'    C8 = password input
'    C10 = message
'    C12 = current user
'    C13 = current role
'
'  USERS SHEET: USERS کاربران (VeryHidden)
'    Table TBL_USERS (or DIM_USER) with CODE,USERNAME,PASSWORD,ROLE,
'    DISPLAY_NAME,ACTIVE
'
'  FLOW:
'    Workbook_Open / Auto_Open -> HideAllExceptLogin
'    QC_Login -> validate -> ShowSheetsByRole
'    QC_Logout -> HideAllExceptLogin
'    Failed attempts logged to LOG_AUDIT
'=======================================================================

Private Const TBL_QC As String = "tblQC"
Private Const TBL_RULE As String = "RULE_STATION"
Private Const TBL_DISP_OLD As String = "DIM_DISPOSITION"
Private Const TBL_USERS As String = "TBL_USERS"
Private Const TBL_USERS_ALT As String = "DIM_USER"
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

Public Const QC_VERSION As String = "4.0.0-login-writepath"

'-----------------------------------------------------------------------
' Auto_Open - runs when workbook opens if macros enabled
' Hides all sheets except LOGIN (VeryHidden for admin sheets)
' Also called by Workbook_Open in ThisWorkbook
'-----------------------------------------------------------------------
Public Sub Auto_Open()
    On Error Resume Next
    HideAllExceptLogin
    On Error GoTo 0
End Sub

Public Sub Workbook_Open_Handler()
    Auto_Open
End Sub

'-----------------------------------------------------------------------
' HideAllExceptLogin - initial state: only LOGIN visible
'-----------------------------------------------------------------------
Public Sub HideAllExceptLogin()
    Dim ws As Worksheet
    Dim loginName As String
    Dim nm As String
    
    On Error GoTo Fail
    loginName = GetSheetName("LOGIN")
    If Len(loginName) = 0 Then loginName = "LOGIN ورود"
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = loginName Then
            ws.Visible = xlSheetVisible
        Else
            ' VeryHidden so user cannot unhide via UI
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws
    
    ' Clear current user display on LOGIN
    On Error Resume Next
    Dim wsL As Worksheet
    Set wsL = ThisWorkbook.Worksheets(loginName)
    If Not wsL Is Nothing Then
        wsL.Range("C12").Value = "(not logged in)"
        wsL.Range("C13").Value = ""
        wsL.Range("C10").Value = "Please login"
        wsL.Range("C6").Value = ""
        wsL.Range("C8").Value = ""
    End If
    On Error GoTo 0
    Exit Sub
Fail:
    ' silent fail on open
End Sub

'-----------------------------------------------------------------------
' QC_Login - main login routine
' Reads C6=user, C8=pass from LOGIN sheet, validates against USERS
'-----------------------------------------------------------------------
Public Sub QC_Login()
    Dim wsL As Worksheet
    Dim wsU As Worksheet
    Dim userIn As String, passIn As String
    Dim found As Boolean, active As String, role As String, dispName As String
    Dim pwd As String, uname As String
    Dim lo As ListObject
    Dim i As Long
    Dim loginName As String
    
    On Error GoTo Fail
    loginName = GetSheetName("LOGIN")
    If Len(loginName) = 0 Then
        MsgBox "LOGIN sheet not found", vbExclamation
        Exit Sub
    End If
    Set wsL = ThisWorkbook.Worksheets(loginName)
    
    userIn = Trim$(CStr(wsL.Range("C6").Value))
    passIn = Trim$(CStr(wsL.Range("C8").Value))
    
    If Len(userIn) = 0 Or Len(passIn) = 0 Then
        wsL.Range("C10").Value = MsgFallback("E_LOGIN", "Invalid login - empty")
        Beep
        MsgBox MsgFallback("E_LOGIN", "Username or password wrong"), vbExclamation, "Login"
        LogLine "LOGIN", "LOGIN_FAIL", "USER", "", userIn & " | empty"
        Exit Sub
    End If
    
    ' Find USERS sheet
    Dim usersSheetName As String
    usersSheetName = GetSheetName("USERS")
    If Len(usersSheetName) = 0 Then usersSheetName = "USERS کاربران"
    
    On Error Resume Next
    Set wsU = ThisWorkbook.Worksheets(usersSheetName)
    On Error GoTo Fail
    
    If wsU Is Nothing Then
        ' fallback to Master DIM_USER
        Set wsU = GetMasterSheet()
        If wsU Is Nothing Then
            MsgBox MsgFallback("E_CONFIG", "Config missing") & " USERS", vbExclamation
            Exit Sub
        End If
    End If
    
    ' Try TBL_USERS first, then DIM_USER
    Set lo = Nothing
    On Error Resume Next
    Set lo = wsU.ListObjects(TBL_USERS)
    If lo Is Nothing Then Set lo = wsU.ListObjects(TBL_USERS_ALT)
    ' If still nothing and wsU is Master, try Master DIM_USER
    If lo Is Nothing Then
        Dim wsM As Worksheet
        Set wsM = GetMasterSheet()
        If Not wsM Is Nothing Then
            Set lo = wsM.ListObjects(TBL_USERS_ALT)
        End If
    End If
    On Error GoTo Fail
    
    If lo Is Nothing Then
        MsgBox "Users table not found (TBL_USERS / DIM_USER)", vbExclamation
        Exit Sub
    End If
    
    found = False
    For i = 1 To lo.ListRows.Count
        uname = Trim$(CStr(lo.ListColumns("USERNAME").DataBodyRange.Cells(i, 1).Value))
        If LCase$(uname) = LCase$(userIn) Then
            pwd = Trim$(CStr(lo.ListColumns("PASSWORD").DataBodyRange.Cells(i, 1).Value))
            role = UCase$(Trim$(CStr(lo.ListColumns("ROLE").DataBodyRange.Cells(i, 1).Value)))
            active = Trim$(CStr(lo.ListColumns("ACTIVE").DataBodyRange.Cells(i, 1).Value))
            dispName = Trim$(CStr(lo.ListColumns("DISPLAY_NAME").DataBodyRange.Cells(i, 1).Value))
            If Len(dispName) = 0 Then dispName = uname
            found = True
            Exit For
        End If
    Next i
    
    If Not found Then
        wsL.Range("C10").Value = MsgFallback("E_LOGIN", "Invalid username or password")
        Beep
        MsgBox MsgFallback("E_LOGIN", "Invalid username or password"), vbExclamation, "Login"
        LogLine "LOGIN", "LOGIN_FAIL", "USER", "", userIn & " | not found"
        Exit Sub
    End If
    
    ' Check active
    If IsInactive(active) Then
        wsL.Range("C10").Value = MsgFallback("E_LOGIN_INACTIVE", "User inactive")
        MsgBox MsgFallback("E_LOGIN_INACTIVE", "User inactive"), vbExclamation, "Login"
        LogLine userIn, "LOGIN_FAIL", "INACTIVE", "", role
        Exit Sub
    End If
    
    ' Check password (case sensitive)
    If pwd <> passIn Then
        wsL.Range("C10").Value = MsgFallback("E_LOGIN", "Invalid username or password")
        Beep
        MsgBox MsgFallback("E_LOGIN", "Invalid username or password"), vbExclamation, "Login"
        LogLine userIn, "LOGIN_FAIL", "PASSWORD", "", "wrong password"
        Exit Sub
    End If
    
    ' Success
    wsL.Range("C12").Value = userIn & " (" & dispName & ")"
    wsL.Range("C13").Value = role
    
    Dim msgKey As String
    Select Case role
        Case "ADMIN"
            msgKey = "OK_LOGIN_ADMIN"
        Case "OPERATOR"
            msgKey = "OK_LOGIN_OPERATOR"
        Case "VIEWER"
            msgKey = "OK_LOGIN_VIEWER"
        Case Else
            msgKey = "OK_LOGIN_VIEWER"
    End Select
    
    wsL.Range("C10").Value = MsgFallback(msgKey, "Login OK - " & role)
    
    ShowSheetsByRole role
    
    LogLine userIn, "LOGIN_OK", "ROLE", "", role
    
    MsgBox MsgFallback(msgKey, "Login OK - " & role) & vbCrLf & userIn & " / " & role, vbInformation, "QC-F-14 Login"
    Exit Sub
Fail:
    FailNow "Login error " & Err.Number & ": " & Err.Description
End Sub

'-----------------------------------------------------------------------
' QC_Logout - hide all except LOGIN
'-----------------------------------------------------------------------
Public Sub QC_Logout()
    Dim wsL As Worksheet
    Dim loginName As String
    Dim curUser As String
    
    On Error GoTo Fail
    loginName = GetSheetName("LOGIN")
    If Len(loginName) = 0 Then Exit Sub
    Set wsL = ThisWorkbook.Worksheets(loginName)
    
    curUser = Trim$(CStr(wsL.Range("C12").Value))
    If Len(curUser) = 0 Or curUser = "(not logged in)" Or InStr(curUser, "not logged") > 0 Then
        MsgBox MsgFallback("E_LOGOUT", "Not logged in"), vbExclamation, "Logout"
        Exit Sub
    End If
    
    LogLine curUser, "LOGOUT", "USER", "", ""
    
    HideAllExceptLogin
    
    wsL.Range("C10").Value = MsgFallback("OK_LOGOUT", "Logged out - only LOGIN visible")
    MsgBox MsgFallback("OK_LOGOUT", "Logged out"), vbInformation, "QC-F-14"
    Exit Sub
Fail:
    FailNow "Logout error " & Err.Number & ": " & Err.Description
End Sub

Private Function IsInactive(ByVal activeVal As String) As Boolean
    Dim v As String
    v = Trim$(activeVal)
    If Len(v) = 0 Then IsInactive = False: Exit Function
    v = LCase$(v)
    If v = "no" Or v = "false" Or v = "0" Or v = "inactive" Or v = "خیر" Then
        IsInactive = True
    Else
        IsInactive = False
    End If
End Function

Private Sub ShowSheetsByRole(ByVal role As String)
    Dim ws As Worksheet
    Dim dict As Object
    Dim loginName As String, formName As String, stagingName As String
    Dim dataName As String, masterName As String, reportsName As String
    Dim dashName As String, dqName As String, setName As String
    Dim helpName As String, logName As String, usersName As String
    Dim helperName As String, verifyName As String
    
    loginName = GetSheetName("LOGIN")
    formName = GetSheetName("FORM")
    stagingName = GetSheetName("STAGING")
    dataName = GetSheetName("DATA")
    masterName = GetSheetName("MASTER")
    reportsName = GetSheetName("REPORTS")
    dashName = GetSheetName("DASH")
    dqName = GetSheetName("DQ")
    setName = GetSheetName("SETTINGS")
    helpName = GetSheetName("HELP")
    logName = GetSheetName("LOG")
    usersName = GetSheetName("USERS")
    helperName = GetSheetName("HELPER")
    verifyName = GetSheetName("VERIFY")
    
    role = UCase$(Trim$(role))
    
    For Each ws In ThisWorkbook.Worksheets
        Select Case role
            Case "ADMIN"
                ' ADMIN sees all
                ws.Visible = xlSheetVisible
                If ws.Name = helperName Or ws.Name = verifyName Then
                    ' keep helper hidden but admin can unhide if needed? Make visible for admin
                    ws.Visible = xlSheetVisible
                End If
            Case "OPERATOR"
                If ws.Name = loginName Or ws.Name = formName Or ws.Name = stagingName Or ws.Name = helpName Then
                    ws.Visible = xlSheetVisible
                Else
                    ws.Visible = xlSheetVeryHidden
                End If
            Case "VIEWER"
                If ws.Name = loginName Or ws.Name = dashName Or ws.Name = reportsName Or ws.Name = dqName Or ws.Name = helpName Then
                    ws.Visible = xlSheetVisible
                Else
                    ws.Visible = xlSheetVeryHidden
                End If
            Case Else
                ' unknown role -> only LOGIN
                If ws.Name = loginName Then
                    ws.Visible = xlSheetVisible
                Else
                    ws.Visible = xlSheetVeryHidden
                End If
        End Select
    Next ws
    
    ' Ensure LOGIN always visible and active
    On Error Resume Next
    ThisWorkbook.Worksheets(loginName).Visible = xlSheetVisible
    ThisWorkbook.Worksheets(loginName).Activate
    On Error GoTo 0
End Sub

Private Function GetSheetName(ByVal key As String) As String
    Dim s As String
    On Error Resume Next
    s = Cfg(CFG_SHEETS, key)
    On Error GoTo 0
    If Len(s) = 0 Then
        ' fallback hardcoded
        Select Case UCase$(key)
            Case "LOGIN": s = "LOGIN ورود"
            Case "FORM": s = "FORM فرم ثبت"
            Case "DATA": s = "Data داده"
            Case "STAGING": s = "STAGING ورود اضطراری"
            Case "MASTER": s = "Master اطلاعات پایه"
            Case "HELPER": s = "Helper محاسبات"
            Case "REPORTS": s = "Reports گزارش‌ها"
            Case "DASH": s = "Dashboard داشبورد"
            Case "DQ": s = "DQ کیفیت داده"
            Case "SETTINGS": s = "Settings تنظیمات"
            Case "HELP": s = "Help راهنما"
            Case "LOG": s = "LOG_AUDIT"
            Case "USERS": s = "USERS کاربران"
            Case "VERIFY": s = "Verify"
        End Select
    End If
    GetSheetName = s
End Function

Private Function GetMasterSheet() As Worksheet
    Dim nm As String
    nm = GetSheetName("MASTER")
    On Error Resume Next
    Set GetMasterSheet = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
End Function

'=======================================================================
' WRITE PATH (from v3) - unchanged logic, but UserName now uses logged user
'=======================================================================
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
        FailNow MsgFallback("E_INCOMPLETE", "Incomplete") & vbCrLf & CellTxt(wsF, "C28")
        Exit Sub
    End If

    If Len(pickBarcode) <> BC_LEN Or Not AllDigits(pickBarcode) Then FailNow MsgFallback("E_BARCODE", "Barcode must be 19 digits"): Exit Sub

    Dim gi As Variant
    gi = Application.Match(pickDate, GridRange("CODE_J_TEXT"), 0)
    If IsError(gi) Then FailNow MsgFallback("E_DATE", "Invalid J date"): Exit Sub

    dups = Application.WorksheetFunction.CountIf(QcCol(wsD, "BARCODE"), pickBarcode)

    If InStr(1, pickStationName, "Return", vbTextCompare) > 0 Or codeStation = "ST-08" Then
        recType = GridVal(GRID_RECTYPE, RT_RETURN)
    ElseIf dups > 0 Then
        recType = GridVal(GRID_RECTYPE, RT_REINSPECTION)
    Else
        recType = GridVal(GRID_RECTYPE, RT_INSPECTION)
    End If

    Dim resCode As String, resName As String
    resCode = PickCode(pickResult)
    resName = PickName(pickResult)
    If Len(resName) = 0 Then resName = resCode
    If UCase$(resName) = "OK" And Len(codeDefect) > 0 Then
        FailNow "RULE-01: OK with defect not allowed"
        Exit Sub
    End If
    If UCase$(resName) = "NOK" And Len(codeDefect) = 0 Then
        FailNow "RULE-02: NOK without defect not allowed"
        Exit Sub
    End If

    Dim dispCode As String, dispName As String
    dispCode = PickCode(pickDisp)
    dispName = PickName(pickDisp)
    If Len(dispName) = 0 Then dispName = dispCode

    If Len(PickCode(pickStatus)) > 0 Then
        curStatus = PickName(pickStatus)
        If Len(curStatus) = 0 Then curStatus = PickCode(pickStatus)
    Else
        If dispName = "REWORK" Then
            curStatus = GridVal(GRID_STATUS, STAT_UNDER)
        ElseIf dispName = "HOLD" Then
            curStatus = GridVal(GRID_STATUS, STAT_OPEN)
        Else
            curStatus = GridVal(GRID_STATUS, STAT_CLOSED)
        End If
    End If

    etLegacy = DispRuleOld(codeStation)
    If Len(etLegacy) = 0 Then etLegacy = StationRule(codeStation)
    If Len(etLegacy) = 0 Then etLegacy = "اصلاحی"

    Dim rowStat As String
    rowStat = GridVal(GRID_STAT, ST_OK)
    If dups > 0 Then rowStat = GridVal(GRID_STAT, ST_DUP)

    r = NextDataRow(wsD)
    If r = 0 Then FailNow MsgFallback("E_CAPACITY", "Capacity full"): Exit Sub

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
    ok = MsgFallback("OK_APPEND", "Appended") & " " & QcText(wsD, "RECORD_ID", r) & " (" & recType & ")"
    If dups > 0 Then ok = ok & vbCrLf & MsgFallback("E_DUP_ACT", "Duplicate flagged")
    MsgBox ok, vbInformation, "QC-F-14 v4"
    Exit Sub
Fail:
    FailNow "QC-ERR " & Err.Number & ": " & Err.Description & " [nothing written]"
End Sub

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
                FailNow MsgFallback("E_INCOMPLETE", "Incomplete") & " row " & r: Exit Sub
            End If
            If Len(pickBarcode) <> BC_LEN Or Not AllDigits(pickBarcode) Then FailNow MsgFallback("E_BARCODE", "Barcode 19") & " row " & r: Exit Sub
            gi = Application.Match(pickDate, GridRange("CODE_J_TEXT"), 0)
            If IsError(gi) Then FailNow MsgFallback("E_DATE", "Invalid date") & " row " & r: Exit Sub

            dups = Application.WorksheetFunction.CountIf(QcCol(wsD, "BARCODE"), pickBarcode)
            If InStr(1, pickStationName, "Return", vbTextCompare) > 0 Then
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
            If tgt = 0 Then FailNow MsgFallback("E_CAPACITY", "Capacity"): Exit Sub

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
    If n = 0 Then MsgBox MsgFallback("E_STAGE_NONE", "No rows ready"), vbExclamation, "QC-F-14": Exit Sub
    LogLine "STAGING", "IMPORT", "ROWS", "", CStr(n)
    MsgBox MsgFallback("OK_IMPORT", "Import OK") & " (" & n & ")", vbInformation, "QC-F-14"
    Exit Sub
Fail:
    FailNow "QC-ERR " & Err.Number & ": " & Err.Description & " [import stopped]"
End Sub

Public Sub QC_VoidRow()
    Dim wsD As Worksheet, lo As ListObject, r As Long, reason As String
    On Error GoTo Fail
    Set wsD = SheetAt("DATA")
    Set lo = wsD.ListObjects(TBL_QC)
    If Selection.Rows.Count <> 1 Or Selection.Areas.Count <> 1 Or Intersect(Selection, lo.DataBodyRange) Is Nothing Then
        FailNow MsgFallback("E_VOID", "Select a Data row"): Exit Sub
    End If
    r = Selection.Row
    reason = InputBox("Enter void reason (required):", "Void Reason")
    If Len(Trim$(reason)) = 0 Then MsgBox "Void without reason not allowed (RULE-07)", vbExclamation: Exit Sub
    PutText wsD, r, "CURRENT_STATUS", GridVal(GRID_STATUS, STAT_VOID)
    PutText wsD, r, "VOID_FLAG", "TRUE"
    PutText wsD, r, "VOID_REASON", reason
    PutText wsD, r, "ROW_STATUS", GridVal(GRID_STAT, ST_VOID)
    LogLine QcText(wsD, "RECORD_ID", r), "VOID", "VOID_REASON", "", reason
    MsgBox MsgFallback("OK_VOID", "Void OK"), vbInformation, "QC-F-14"
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
    MsgBox "QC-F-14 " & QC_VERSION & vbCrLf & MsgFallback("ABOUT", "QC-F-14 v4 login + writepath") & vbCrLf & "rows: " & n, vbInformation, "QC-F-14"
End Sub

'--- write full v4 row ---
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
    PutText wsD, r, "ENTRY_CHANNEL", LabelFallback("CHANNEL_FORM", "FORM-MACRO")
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

Private Function MsgFallback(ByVal k As String, ByVal fb As String) As String
    Dim s As String
    On Error Resume Next
    s = Cfg(CFG_MSG, k)
    On Error GoTo 0
    If Len(s) = 0 Then s = fb
    MsgFallback = s
End Function

Private Function LabelFallback(ByVal k As String, ByVal fb As String) As String
    LabelFallback = MsgFallback(k, fb)
End Function

Private Function SheetAt(ByVal k As String) As Worksheet
    Dim s As String
    s = GetSheetName(k)
    If Len(s) = 0 Then Err.Raise vbObjectError + 522, "QC", "Config missing (" & k & ")"
    On Error Resume Next
    Set SheetAt = ThisWorkbook.Worksheets(s)
    On Error GoTo 0
    If SheetAt Is Nothing Then Err.Raise vbObjectError + 523, "QC", "Sheet not found (" & s & ")"
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
    Set ws = GetMasterSheet()
    If ws Is Nothing Then Exit Function
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
    Set ws = GetMasterSheet()
    If ws Is Nothing Then Exit Function
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
    MsgBox s, vbExclamation, "QC-F-14 v4 -- nothing written"
    Err.Clear
End Sub

Private Function UserName() As String
    Dim wsL As Worksheet
    Dim loginName As String
    Dim cur As String
    On Error Resume Next
    loginName = GetSheetName("LOGIN")
    Set wsL = ThisWorkbook.Worksheets(loginName)
    If Not wsL Is Nothing Then
        cur = Trim$(CStr(wsL.Range("C12").Value))
        If Len(cur) > 0 And cur <> "(not logged in)" And InStr(cur, "not logged") = 0 Then
            ' extract username before space or (
            Dim p As Long
            p = InStr(cur, " ")
            If p > 0 Then cur = Left$(cur, p - 1)
            p = InStr(cur, "(")
            If p > 0 Then cur = Left$(cur, p - 1)
            cur = Trim$(cur)
            If Len(cur) > 0 Then
                UserName = cur
                Exit Function
            End If
        End If
    End If
    On Error GoTo 0
    UserName = Environ$("USERNAME")
    If Len(UserName) = 0 Then UserName = Environ$("USER")
    If Len(UserName) = 0 Then UserName = "MIGRATION"
End Function
