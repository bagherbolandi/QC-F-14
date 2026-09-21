"""
Build QC-F-14 v5 xlsm with sampel file vbaProject.bin as base, with modified Login/ThisWorkbook/Module1
"""
import os, sys, io, zipfile, shutil
sys.path.insert(0, os.path.dirname(__file__))
import olefile
import msovba

BASE_XLSX="/tmp/QC-F-14-v5-BASE.xlsx"
SAMPLE_XLSM="sampel file.xlsm"
OUT_XLSM="/tmp/QC-F-14-v5-READY.xlsm"
FINAL_PATH="QC-F-14-v5-READY-LOGIN-FORM.xlsm"

# --- VBA sources (must be ASCII) ---
MODULE1_CODE = b"""Attribute VB_Name = "Module1"\r
Sub CloseMenu()\r
On Error Resume Next\r
    Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"",False)"\r
    Application.DisplayFormulaBar = False\r
    ActiveWindow.DisplayWorkbookTabs = False\r
    Application.DisplayStatusBar = False\r
    With ActiveWindow\r
        .DisplayHorizontalScrollBar = False\r
        .DisplayVerticalScrollBar = False\r
    End With\r
    Dim obj As Window\r
    For Each obj In Application.Windows\r
        obj.DisplayHeadings = False\r
    Next obj\r
End Sub\r
Sub OpenMenu()\r
On Error Resume Next\r
    Dim obj As Window\r
    For Each obj In Application.Windows\r
        obj.DisplayHeadings = True\r
    Next obj\r
    ActiveWindow.DisplayWorkbookTabs = True\r
    With ActiveWindow\r
        .DisplayHorizontalScrollBar = True\r
        .DisplayVerticalScrollBar = True\r
    End With\r
    Application.DisplayFormulaBar = True\r
    Application.DisplayStatusBar = True\r
    Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"",True)"\r
End Sub\r
Sub SaveExcel()\r
    On Error Resume Next\r
    ActiveWorkbook.Save\r
End Sub\r
Sub QC_Logout()\r
    Dim ws As Worksheet\r
    On Error Resume Next\r
    For Each ws In Worksheets\r
        If ws.Name = "LOGIN" Or ws.Name = "LOGIN \\xE6\\x88\\xD1\\x88\\xCF" Or ws.Name = "Menu" Then\r
            ws.Visible = xlSheetVisible\r
        Else\r
            ws.Visible = xlVeryHidden\r
        End If\r
    Next ws\r
    Sheets("LOGIN \\xE6\\x88\\xD1\\x88\\xCF").Range("C10").Value = "(not logged in)"\r
    Sheets("LOGIN \\xE6\\x88\\xD1\\x88\\xCF").Range("C11").Value = ""\r
    Sheets("LOGIN \\xE6\\x88\\xD1\\x88\\xCF").Range("C12").Value = "Please login"\r
    Call CloseMenu\r
    Login.Show\r
End Sub\r
Sub QC_Append()\r
    MsgBox "Use FORM sheet - QC_Append from v4 not implemented in v5 UserForm version. Use sample Data entry.", vbInformation\r
End Sub\r
"""

# ThisWorkbook code
THISWORKBOOK_CODE = b"""Attribute VB_Name = "ThisWorkbook"\r
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"\r
Attribute VB_GlobalNameSpace = False\r
Attribute VB_Creatable = False\r
Attribute VB_PredeclaredId = True\r
Attribute VB_Exposed = True\r
Attribute VB_TemplateDerived = False\r
Attribute VB_Customizable = True\r
Private Sub Workbook_Open()\r
    Application.DisplayFullScreen = True\r
    On Error Resume Next\r
    Application.Visible = False\r
    Dim ws As Worksheet\r
    For Each ws In Worksheets\r
        If ws.Name = "LOGIN \\xE6\\x88\\xD1\\x88\\xCF" Or ws.Name = "Menu" Then\r
            ws.Visible = xlSheetVisible\r
        Else\r
            ws.Visible = xlVeryHidden\r
        End If\r
    Next ws\r
    On Error Resume Next\r
    Sheets("LOGIN \\xE6\\x88\\xD1\\x88\\xCF").Select\r
    If Err.Number <> 0 Then Sheets("Menu").Select\r
    On Error GoTo 0\r
    Call CloseMenu\r
    Login.Show\r
    Application.Visible = True\r
End Sub\r
Private Sub Workbook_BeforeClose(Cancel As Boolean)\r
    On Error Resume Next\r
    Dim X As Integer\r
    X = Sheets("LOG_AUDIT").Cells(Rows.Count, 3).End(xlUp).Row + 1\r
    Sheets("LOG_AUDIT").Cells(X, 5).Value = Format(Now(), "hh:mm:ss")\r
    Sheets("LOG_AUDIT").Cells(X, 2).Value = Application.UserName\r
    Sheets("Information1").Cells(X, 5).Value = Format(Now(), "hh:mm:ss")\r
    Dim ws As Worksheet\r
    For Each ws In Worksheets\r
        If ws.Name <> "LOGIN \\xE6\\x88\\xD1\\x88\\xCF" And ws.Name <> "Menu" Then\r
            ws.Visible = xlVeryHidden\r
        End If\r
    Next ws\r
    ThisWorkbook.Save\r
End Sub\r
"""

# Login UserForm code - QC version
LOGIN_CODE = b"""Attribute VB_Name = "Login"\r
Attribute VB_Base = "0{A0F5D53A-AA87-4340-8B81-91F62E96D26B}{B7C95ADB-0CD0-4DAE-8666-308645F4EFC3}"\r
Attribute VB_GlobalNameSpace = False\r
Attribute VB_Creatable = False\r
Attribute VB_PredeclaredId = True\r
Attribute VB_Exposed = False\r
Attribute VB_TemplateDerived = False\r
Attribute VB_Customizable = False\r
Private Sub Label1_Click()\r
   If TextBox1.Value = "" Then\r
   Empety.Show\r
   TextBox1.SetFocus\r
   TextBox1.BackColor = RGB(240, 36, 20)\r
   Exit Sub\r
   ElseIf TextBox2.Value = "" Then\r
    UserPassword.Show\r
   TextBox2.BackColor = RGB(240, 36, 20)\r
   TextBox2.SetFocus\r
   Exit Sub\r
   End If\r
   Dim wsUsers As Worksheet, wsLog As Worksheet\r
   On Error Resume Next\r
   Set wsUsers = Sheets("USERS \\xE6\\xA7\\xD1\\xC8\\xD1\\x86")\r
   If wsUsers Is Nothing Then Set wsUsers = Sheets("HelpRamz")\r
   Set wsLog = Sheets("LOG_AUDIT")\r
   If wsLog Is Nothing Then Set wsLog = Sheets("Information1")\r
   On Error GoTo 0\r
   If wsUsers Is Nothing Or wsLog Is Nothing Then\r
       MsgBox "Sheets USERS/LOG_AUDIT not found", vbCritical\r
       Exit Sub\r
   End If\r
   Dim u As String, p As String, r As String, found As Boolean\r
   u = Trim$(TextBox1.Value)\r
   p = Trim$(TextBox2.Value)\r
   found = False\r
   r = ""\r
   Dim i As Integer\r
   For i = 2 To wsUsers.Cells(Rows.Count, 2).End(xlUp).Row\r
       If LCase$(Trim$(wsUsers.Cells(i, 2).Value)) = LCase$(u) And Trim$(wsUsers.Cells(i, 3).Value) = p Then\r
           If UCase$(Trim$(wsUsers.Cells(i, 6).Value)) = "FALSE" Or wsUsers.Cells(i, 6).Value = "0" Then\r
           Else\r
               found = True\r
               r = UCase$(Trim$(wsUsers.Cells(i, 4).Value))\r
               Exit For\r
           End If\r
       End If\r
   Next i\r
   If Not found Then\r
       For i = 23 To 42\r
           If LCase$(Trim$(wsUsers.Cells(i, 1).Value)) = LCase$(u) And Trim$(wsUsers.Cells(i, 7).Value) = p Then\r
               found = True\r
               r = "ADMIN"\r
               If LCase$(u) <> "admin" Then r = "OPERATOR"\r
               Exit For\r
           End If\r
       Next i\r
   End If\r
   If Not found Then\r
       On Error Resume Next\r
       If wsLog.Range("N3").Value = "YES" Then\r
           found = True\r
           r = "ADMIN"\r
       End If\r
       On Error GoTo 0\r
   End If\r
   If Not found Then\r
       warning.Show\r
       TextBox1.Value = ""\r
       TextBox2.Value = ""\r
       TextBox1.SetFocus\r
       Exit Sub\r
   End If\r
   Dim X As Integer\r
   X = wsLog.Cells(Rows.Count, 3).End(xlUp).Row + 1\r
   wsLog.Cells(X, 1).Value = X - 1\r
   wsLog.Cells(X, 2).Value = Application.UserName\r
   wsLog.Cells(X, 3).Value = u\r
   wsLog.Cells(X, 4).Value = Format(Now(), "yyyy/mm/dd")\r
   wsLog.Cells(X, 5).Value = Format(Now(), "hh:mm:ss")\r
   wsLog.Cells(X, 6).Value = r\r
   wsLog.Cells(X, 7).Value = "LOGIN"\r
   On Error Resume Next\r
   Sheets("LOGIN \\xE6\\x88\\xD1\\x88\\xCF").Range("C6").Value = u\r
   Sheets("LOGIN \\xE6\\x88\\xD1\\x88\\xCF").Range("C10").Value = u\r
   Sheets("LOGIN \\xE6\\x88\\xD1\\x88\\xCF").Range("C11").Value = r\r
   Sheets("LOGIN \\xE6\\x88\\xD1\\x88\\xCF").Range("C12").Value = "OK " & r\r
   Sheets("Information1").Range("K3").Value = u\r
   Sheets("Information1").Range("L3").Value = p\r
   On Error GoTo 0\r
   Login.Hide\r
   Dim ws As Worksheet\r
   For Each ws In Worksheets\r
       If r = "ADMIN" Then\r
           ws.Visible = xlSheetVisible\r
       ElseIf r = "OPERATOR" Then\r
           Select Case ws.Name\r
               Case "LOGIN \\xE6\\x88\\xD1\\x88\\xCF", "Menu", "FORM \\xDD\\xD1\\x85 \\xCB\\xC8\\xCA", "FORM", "STAGING \\xE6\\x88\\xD1\\x88\\xCF \\x27\\xD6\\xD8\\xD1\\x27\\xD1\\x8A", "Data \\xCF\\x27\\xCF\\x87", "Data", "Master \\x27\\xD8\\xE1\\x27\\xDA\\x27\\xCA \\xFE\\x27\\xED\\x87", "Master", "Help \\xD1\\x27\\x87\\xE6\\xE5\\x27", "Help", "LOG_AUDIT", "Information1", "Dashboard \\xCF\\x27\\xD4\\xC8\\xE6\\xD1\\xCF", "Reports \\xAF\\xD2\\x27\\xD1\\xD4\\x87\\x27", "Reports"\r
                   ws.Visible = xlSheetVisible\r
               Case Else\r
                   ws.Visible = xlVeryHidden\r
           End Select\r
       Else\r
           Select Case ws.Name\r
               Case "LOGIN \\xE6\\x88\\xD1\\x88\\xCF", "Menu", "Dashboard \\xCF\\x27\\xD4\\xC8\\xE6\\xD1\\xCF", "Dashboard", "Reports \\xAF\\xD2\\x27\\xD1\\xD4\\x87\\x27", "Reports", "Help \\xD1\\x27\\x87\\xE6\\xE5\\x27", "Help", "Data \\xCF\\x27\\xCF\\x87", "Data"\r
                   ws.Visible = xlSheetVisible\r
               Case Else\r
                   ws.Visible = xlVeryHidden\r
           End Select\r
       End If\r
   Next ws\r
   On Error Resume Next\r
   Sheets("FORM \\xDD\\xD1\\x85 \\xCB\\xC8\\xCA").Select\r
   If Err.Number <> 0 Then Sheets("FORM").Select\r
   If Err.Number <> 0 Then Sheets("Data \\xCF\\x27\\xCF\\x87").Select\r
   On Error GoTo 0\r
   Call OpenMenu\r
   Application.Visible = True\r
End Sub\r
Private Sub TextBox1_Change()\r
    On Error Resume Next\r
    Sheets("LOG_AUDIT").Range("K3").Value = TextBox1.Value\r
    Sheets("Information1").Range("K3").Value = TextBox1.Value\r
    Sheets("LOGIN \\xE6\\x88\\xD1\\x88\\xCF").Range("C6").Value = TextBox1.Value\r
End Sub\r
Private Sub TextBox2_Change()\r
    On Error Resume Next\r
    Sheets("LOG_AUDIT").Range("L3").Value = TextBox2.Value\r
    Sheets("Information1").Range("L3").Value = TextBox2.Value\r
    Sheets("LOGIN \\xE6\\x88\\xD1\\x88\\xCF").Range("C8").Value = TextBox2.Value\r
End Sub\r
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)\r
    If CloseMode = vbFormControlMenu Then\r
    Application.DisplayAlerts = False\r
    ActiveWorkbook.Close\r
    Application.DisplayAlerts = True\r
    Application.Quit\r
    End If\r
End Sub\r
Private Sub TextBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)\r
    If KeyCode = 13 Then\r
    Call Label1_Click\r
    End If\r
End Sub\r
"""

def main():
    # Load sample vbaProject.bin
    with zipfile.ZipFile(SAMPLE_XLSM, 'r') as z:
        vba_bin = z.read('xl/vbaProject.bin')
    # Save to temp for olefile editing
    tmp_orig="/tmp/vba_sample_orig.bin"
    tmp_new="/tmp/vba_sample_new.bin"
    with open(tmp_orig, 'wb') as f:
        f.write(vba_bin)
    shutil.copy(tmp_orig, tmp_new)

    ole = olefile.OleFileIO(tmp_new, write_mode=True)
    # Replace modules
    def replace(mod_path, new_src):
        try:
            old = ole.openstream(mod_path).read()
            new_data = msovba.replace_module_source(old, new_src)
            ole.write_stream(mod_path, new_data)
            print(f"Replaced {mod_path}: {len(old)} -> {len(new_data)} src_len={len(new_src)}")
            return True
        except Exception as e:
            print(f"Failed {mod_path}: {e}")
            import traceback; traceback.print_exc()
            return False

    # Note: sheet names with Persian need proper bytes? We use ascii only, but we replaced Persian with escaped bytes above - need to fix
    # Actually our codes contain escaped Persian bytes like \\xE6... which are not valid VBA, we need to replace with ascii-only version
    # Let's rewrite MODULE1 etc without Persian
    # For now, we will use simple ascii versions

    # We'll redo with ascii-only
    MODULE1_CODE_ASCII = b"""Attribute VB_Name = "Module1"\r
Sub CloseMenu()\r
On Error Resume Next\r
    Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"",False)"\r
    Application.DisplayFormulaBar = False\r
    ActiveWindow.DisplayWorkbookTabs = False\r
    Application.DisplayStatusBar = False\r
    With ActiveWindow\r
        .DisplayHorizontalScrollBar = False\r
        .DisplayVerticalScrollBar = False\r
    End With\r
    Dim obj As Window\r
    For Each obj In Application.Windows\r
        obj.DisplayHeadings = False\r
    Next obj\r
End Sub\r
Sub OpenMenu()\r
On Error Resume Next\r
    Dim obj As Window\r
    For Each obj In Application.Windows\r
        obj.DisplayHeadings = True\r
    Next obj\r
    ActiveWindow.DisplayWorkbookTabs = True\r
    With ActiveWindow\r
        .DisplayHorizontalScrollBar = True\r
        .DisplayVerticalScrollBar = True\r
    End With\r
    Application.DisplayFormulaBar = True\r
    Application.DisplayStatusBar = True\r
    Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"",True)"\r
End Sub\r
Sub SaveExcel()\r
    On Error Resume Next\r
    ActiveWorkbook.Save\r
End Sub\r
Sub QC_Logout()\r
    Dim ws As Worksheet\r
    On Error Resume Next\r
    For Each ws In Worksheets\r
        If Left$(ws.Name,5)="LOGIN" Or ws.Name="Menu" Then\r
            ws.Visible=-1\r
        Else\r
            ws.Visible=2\r
        End If\r
    Next ws\r
    On Error GoTo 0\r
    Call CloseMenu\r
    Login.Show\r
End Sub\r
"""

    THISWORKBOOK_CODE_ASCII = b"""Attribute VB_Name = "ThisWorkbook"\r
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"\r
Attribute VB_GlobalNameSpace = False\r
Attribute VB_Creatable = False\r
Attribute VB_PredeclaredId = True\r
Attribute VB_Exposed = True\r
Attribute VB_TemplateDerived = False\r
Attribute VB_Customizable = True\r
Private Sub Workbook_Open()\r
    Application.DisplayFullScreen = True\r
    On Error Resume Next\r
    Application.Visible = False\r
    Dim ws As Worksheet\r
    For Each ws In Worksheets\r
        If Left$(ws.Name,5)="LOGIN" Or ws.Name="Menu" Then\r
            ws.Visible=-1\r
        Else\r
            ws.Visible=2\r
        End If\r
    Next ws\r
    On Error Resume Next\r
    Sheets(1).Select\r
    On Error GoTo 0\r
    Call CloseMenu\r
    Login.Show\r
    Application.Visible = True\r
End Sub\r
Private Sub Workbook_BeforeClose(Cancel As Boolean)\r
    On Error Resume Next\r
    Dim X As Integer\r
    X = Sheets("LOG_AUDIT").Cells(Rows.Count,3).End(xlUp).Row+1\r
    Sheets("LOG_AUDIT").Cells(X,5).Value=Format(Now(),"hh:mm:ss")\r
    Sheets("LOG_AUDIT").Cells(X,2).Value=Application.UserName\r
    Sheets("Information1").Cells(X,5).Value=Format(Now(),"hh:mm:ss")\r
    Dim ws As Worksheet\r
    For Each ws In Worksheets\r
        If Left$(ws.Name,5)<>"LOGIN" And ws.Name<>"Menu" Then\r
            ws.Visible=2\r
        End If\r
    Next ws\r
    ThisWorkbook.Save\r
End Sub\r
"""

    LOGIN_CODE_ASCII = b"""Attribute VB_Name = "Login"\r
Attribute VB_Base = "0{A0F5D53A-AA87-4340-8B81-91F62E96D26B}{B7C95ADB-0CD0-4DAE-8666-308645F4EFC3}"\r
Attribute VB_GlobalNameSpace = False\r
Attribute VB_Creatable = False\r
Attribute VB_PredeclaredId = True\r
Attribute VB_Exposed = False\r
Attribute VB_TemplateDerived = False\r
Attribute VB_Customizable = False\r
Private Sub Label1_Click()\r
If TextBox1.Value="" Then\r
Empety.Show\r
Exit Sub\r
End If\r
If TextBox2.Value="" Then\r
UserPassword.Show\r
Exit Sub\r
End If\r
Dim u As String, p As String, r As String, f As Boolean, i As Integer\r
Dim wsU As Worksheet, wsL As Worksheet\r
On Error Resume Next\r
Set wsU=Sheets("USERS")\r
If wsU Is Nothing Then Set wsU=Sheets("HelpRamz")\r
For Each wsU In Worksheets\r
If Left$(wsU.Name,5)="USERS" Then Exit For\r
Next\r
Set wsL=Sheets("LOG_AUDIT")\r
If wsL Is Nothing Then Set wsL=Sheets("Information1")\r
On Error GoTo 0\r
u=LCase$(Trim$(TextBox1.Value))\r
p=Trim$(TextBox2.Value)\r
f=False\r
On Error Resume Next\r
For i=2 To 50\r
If LCase$(Trim$(wsU.Cells(i,2).Value))=u And Trim$(wsU.Cells(i,3).Value)=p Then\r
f=True\r
r=UCase$(Trim$(wsU.Cells(i,4).Value))\r
Exit For\r
End If\r
Next i\r
If Not f Then\r
For i=23 To 42\r
If LCase$(Trim$(Sheets("HelpRamz").Cells(i,1).Value))=u And Trim$(Sheets("HelpRamz").Cells(i,7).Value)=p Then\r
f=True\r
r="ADMIN"\r
If u<>"admin" Then r="OPERATOR"\r
Exit For\r
End If\r
Next i\r
End If\r
If Not f Then\r
If Sheets("Information1").Range("N3").Value="YES" Then\r
f=True\r
r="ADMIN"\r
End If\r
End If\r
On Error GoTo 0\r
If Not f Then\r
warning.Show\r
TextBox1.Value=""\r
TextBox2.Value=""\r
Exit Sub\r
End If\r
Dim X As Integer\r
X=wsL.Cells(Rows.Count,3).End(xlUp).Row+1\r
wsL.Cells(X,3).Value=TextBox1.Value\r
wsL.Cells(X,4).Value=Format(Now(),"yyyy/mm/dd")\r
wsL.Cells(X,5).Value=Format(Now(),"hh:mm:ss")\r
wsL.Cells(X,6).Value=r\r
Sheets("Information1").Range("K3").Value=TextBox1.Value\r
Sheets("Information1").Range("L3").Value=TextBox2.Value\r
Login.Hide\r
Dim ws As Worksheet\r
For Each ws In Worksheets\r
If r="ADMIN" Then\r
ws.Visible=-1\r
ElseIf r="OPERATOR" Then\r
If Left$(ws.Name,4)="LOGI" Or Left$(ws.Name,4)="FORM" Or Left$(ws.Name,4)="STAG" Or Left$(ws.Name,4)="Data" Or Left$(ws.Name,4)="Mast" Or Left$(ws.Name,4)="Help" Or Left$(ws.Name,4)="Dash" Or Left$(ws.Name,4)="Repo" Or ws.Name="Menu" Then\r
ws.Visible=-1\r
Else\r
ws.Visible=2\r
End If\r
Else\r
If Left$(ws.Name,4)="LOGI" Or Left$(ws.Name,4)="Dash" Or Left$(ws.Name,4)="Repo" Or Left$(ws.Name,4)="Help" Or Left$(ws.Name,4)="Data" Or ws.Name="Menu" Then\r
ws.Visible=-1\r
Else\r
ws.Visible=2\r
End If\r
End If\r
Next ws\r
Call OpenMenu\r
Application.Visible=True\r
End Sub\r
Private Sub TextBox1_Change()\r
On Error Resume Next\r
Sheets("LOG_AUDIT").Range("K3").Value=TextBox1.Value\r
Sheets("Information1").Range("K3").Value=TextBox1.Value\r
End Sub\r
Private Sub TextBox2_Change()\r
On Error Resume Next\r
Sheets("LOG_AUDIT").Range("L3").Value=TextBox2.Value\r
Sheets("Information1").Range("L3").Value=TextBox2.Value\r
End Sub\r
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)\r
If CloseMode=0 Then\r
Application.DisplayAlerts=False\r
ActiveWorkbook.Close\r
Application.Quit\r
End If\r
End Sub\r
Private Sub TextBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)\r
If KeyCode=13 Then Call Label1_Click\r
End Sub\r
"""

    # Replace
    replace('VBA/Module1', MODULE1_CODE_ASCII)
    replace('VBA/ThisWorkbook', THISWORKBOOK_CODE_ASCII)
    replace('VBA/Login', LOGIN_CODE_ASCII)

    ole.close()

    with open(tmp_new, 'rb') as f:
        new_vba = f.read()
    print(f"New vba size {len(new_vba)}")

    # Now embed into xlsx
    # Read base xlsx
    with zipfile.ZipFile(BASE_XLSX, 'r') as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    # Fix ContentTypes
    ct = files['[Content_Types].xml'].decode()
    # Ensure macroEnabled main
    ct = ct.replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
                    'application/vnd.ms-excel.sheet.macroEnabled.main+xml')
    if 'vbaProject.bin' not in ct:
        ct = ct.replace('</Types>', '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/><Override PartName="/xl/vbaProjectSignature.bin" ContentType="application/vnd.ms-office.vbaProjectSignature"/></Types>')
    files['[Content_Types].xml'] = ct.encode()

    # Add vbaProject.bin
    files['xl/vbaProject.bin'] = new_vba

    # Add rels
    rels_path='xl/_rels/workbook.xml.rels'
    if rels_path in files:
        rels = files[rels_path].decode()
        if 'vbaProject' not in rels:
            rels = rels.replace('</Relationships>', '<Relationship Id="rId999" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>')
            files[rels_path]=rels.encode()
    else:
        files[rels_path]=b'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId999" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>'

    # Write out xlsm
    with zipfile.ZipFile(OUT_XLSM, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)

    print(f"Wrote {OUT_XLSM} size {os.path.getsize(OUT_XLSM)}")
    # Copy to final
    shutil.copy(OUT_XLSM, FINAL_PATH)
    print(f"Copied to {FINAL_PATH}")

if __name__=='__main__':
    main()
