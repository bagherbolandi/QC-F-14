"""
Build v6 final — complete returned parts software with compact VBA that fits sample's modules
"""
import os, sys, io, zipfile, shutil
sys.path.insert(0, os.path.dirname(__file__))
import olefile
import msovba

BASE_XLSX="/tmp/QC-F-14-v6-BASE.xlsx"
SAMPLE_XLSM="sampel file.xlsm"
OUT_XLSM="/tmp/QC-F-14-v6-READY.xlsm"
FINAL1="QC-F-14-v6-FINAL-RETURN-SOFTWARE.xlsm"
FINAL2="QC-F-14-v6-COMPLETE-SOFTWARE.xlsm"

# --- VBA modules (ASCII only, compact) ---
MODULE1_CODE = b"""Attribute VB_Name = "Module1"\r
Sub CloseMenu()\r
On Error Resume Next\r
Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"",False)"\r
Application.DisplayFormulaBar=False\r
ActiveWindow.DisplayWorkbookTabs=False\r
Application.DisplayStatusBar=False\r
With ActiveWindow\r
.DisplayHorizontalScrollBar=False\r
.DisplayVerticalScrollBar=False\r
End With\r
Dim o As Window\r
For Each o In Application.Windows\r
o.DisplayHeadings=False\r
Next\r
End Sub\r
Sub OpenMenu()\r
On Error Resume Next\r
Dim o As Window\r
For Each o In Application.Windows\r
o.DisplayHeadings=True\r
Next\r
ActiveWindow.DisplayWorkbookTabs=True\r
With ActiveWindow\r
.DisplayHorizontalScrollBar=True\r
.DisplayVerticalScrollBar=True\r
End With\r
Application.DisplayFormulaBar=True\r
Application.DisplayStatusBar=True\r
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
Next\r
Call CloseMenu\r
Login.Show\r
End Sub\r
"""

THISWORKBOOK_CODE = b"""Attribute VB_Name = "ThisWorkbook"\r
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"\r
Attribute VB_GlobalNameSpace = False\r
Attribute VB_Creatable = False\r
Attribute VB_PredeclaredId = True\r
Attribute VB_Exposed = True\r
Attribute VB_TemplateDerived = False\r
Attribute VB_Customizable = True\r
Private Sub Workbook_Open()\r
Application.DisplayFullScreen=True\r
On Error Resume Next\r
Application.Visible=False\r
Dim ws As Worksheet\r
For Each ws In Worksheets\r
If Left$(ws.Name,5)="LOGIN" Or ws.Name="Menu" Then\r
ws.Visible=-1\r
Else\r
ws.Visible=2\r
End If\r
Next\r
Sheets(1).Select\r
Call CloseMenu\r
Login.Show\r
Application.Visible=True\r
End Sub\r
Private Sub Workbook_BeforeClose(Cancel As Boolean)\r
On Error Resume Next\r
Dim X As Integer\r
X=Sheets("LOG_AUDIT").Cells(Rows.Count,3).End(xlUp).Row+1\r
Sheets("LOG_AUDIT").Cells(X,5).Value=Format(Now(),"hh:mm:ss")\r
Sheets("Information1").Cells(X,5).Value=Format(Now(),"hh:mm:ss")\r
Dim ws As Worksheet\r
For Each ws In Worksheets\r
If Left$(ws.Name,5)<>"LOGIN" And ws.Name<>"Menu" Then\r
ws.Visible=2\r
End If\r
Next\r
ThisWorkbook.Save\r
End Sub\r
"""

LOGIN_CODE = b"""Attribute VB_Name = "Login"\r
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
If Left$(ws.Name,4)="LOGI" Or Left$(ws.Name,4)="FORM" Or Left$(ws.Name,4)="REtu" Or Left$(ws.Name,4)="RETU" Or Left$(ws.Name,4)="Data" Or Left$(ws.Name,4)="Mast" Or Left$(ws.Name,4)="Help" Or Left$(ws.Name,4)="Dash" Or Left$(ws.Name,4)="Repo" Or ws.Name="Menu" Then\r
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

MODULE2_CODE = b"""Attribute VB_Name = "Module2"\r
Sub QC_Append_RETURN()\r
Dim F As Worksheet, D As Worksheet, r As Long\r
On Error GoTo EH\r
Set F=Sheets(15)\r
Set D=Sheets(2)\r
If F.Range("C6")="" Then MsgBox "Date": Exit Sub\r
If F.Range("C14")="" Then MsgBox "BC": Exit Sub\r
If Len(F.Range("C14"))<>19 Then MsgBox "19": Exit Sub\r
r=D.Cells(D.Rows.Count,1).End(-4162).Row+1\r
D.Cells(r,1)="QC-" & Format$(r-1,"00000000")\r
D.Cells(r,2)="RETURN"\r
D.Cells(r,3)=F.Range("C6")\r
D.Cells(r,7)=F.Range("C8")\r
D.Cells(r,8)=F.Range("C10")\r
D.Cells(r,12)=F.Range("F6")\r
D.Cells(r,16)=F.Range("C14")\r
D.Cells(r,20)=F.Range("F12")\r
D.Cells(r,22)=F.Range("F14")\r
D.Cells(r,24)=Environ$("USERNAME")\r
D.Cells(r,25)=Now\r
MsgBox "RETURN " & D.Cells(r,1), vbInformation\r
Exit Sub\r
EH:\r
MsgBox Err.Description, vbCritical\r
End Sub\r
Sub QC_Clear_RETURN()\r
On Error Resume Next\r
Sheets(15).Range("C6,C8,C10,F6,C14,F12,F14").ClearContents\r
End Sub\r
"""

MODULE3_CODE = b"""Attribute VB_Name = "Module3"\r
Sub QC_Append()\r
Dim wsF As Worksheet, wsD As Worksheet, r As Long\r
On Error GoTo EH\r
Set wsF=Sheets(3)\r
Set wsD=Sheets(2)\r
If Trim$(wsF.Range("C6").Value)="" Then MsgBox "Date required": Exit Sub\r
If Trim$(wsF.Range("C14").Value)="" Then MsgBox "Barcode required": Exit Sub\r
If Len(Trim$(wsF.Range("C14").Value))<>19 Then MsgBox "Barcode 19": Exit Sub\r
r=wsD.Cells(wsD.Rows.Count,1).End(-4162).Row+1\r
wsD.Cells(r,1).Value="QC-" & Format$(r-1,"00000000")\r
wsD.Cells(r,2).Value="INSPECTION"\r
wsD.Cells(r,3).Value=wsF.Range("C6").Value\r
wsD.Cells(r,7).Value=wsF.Range("C8").Value\r
wsD.Cells(r,8).Value=wsF.Range("C10").Value\r
wsD.Cells(r,10).Value=wsF.Range("C12").Value\r
wsD.Cells(r,12).Value=wsF.Range("F6").Value\r
wsD.Cells(r,14).Value=wsF.Range("F8").Value\r
wsD.Cells(r,16).Value=wsF.Range("C14").Value\r
wsD.Cells(r,19).Value=wsF.Range("F10").Value\r
wsD.Cells(r,20).Value=wsF.Range("F12").Value\r
wsD.Cells(r,22).Value=wsF.Range("F14").Value\r
wsD.Cells(r,23).Value=wsF.Range("C16").Value\r
wsD.Cells(r,28).Value=wsF.Range("F16").Value\r
wsD.Cells(r,24).Value=Environ$("USERNAME")\r
wsD.Cells(r,25).Value=Now\r
MsgBox "Recorded " & wsD.Cells(r,1).Value, vbInformation\r
Exit Sub\r
EH:\r
MsgBox "Error " & Err.Description, vbCritical\r
End Sub\r
Sub QC_ClearForm()\r
On Error Resume Next\r
Sheets(3).Range("C6,C8,C10,C12,F6,F8,C14,F10,F12,F14,C16,F16").ClearContents\r
End Sub\r
"""

MODULE4_CODE = b"""Attribute VB_Name = "Module4"\r
Sub QC_VoidRow()\r
Dim wsD As Worksheet, r As Long, rs As String\r
On Error GoTo EH\r
Set wsD=Sheets(2)\r
r=ActiveCell.Row\r
If r<2 Then MsgBox "Select data row": Exit Sub\r
rs=InputBox("Void reason:","Void")\r
If Trim$(rs)="" Then MsgBox "Reason required": Exit Sub\r
wsD.Cells(r,23).Value="VOID"\r
wsD.Cells(r,29).Value="TRUE"\r
wsD.Cells(r,30).Value=rs\r
MsgBox "Void OK", vbInformation\r
Exit Sub\r
EH:\r
MsgBox "Error " & Err.Description, vbCritical\r
End Sub\r
Sub QC_About()\r
MsgBox "QC-F-14 v6 - Returned Parts Software" & vbCrLf & "Login: admin/admin123" & vbCrLf & "Operator: operator/1234", vbInformation\r
End Sub\r
Sub QC_Import_Staging()\r
MsgBox "Use STAGING sheet - copy ready rows to Data", vbInformation\r
End Sub\r
"""

MODULE5_CODE = b"""Attribute VB_Name = "Module5"\r
Sub ShowDashboard()\r
On Error Resume Next\r
Sheets("Dashboard \\xCF\\x27\\xD4\\xC8\\xE6\\xD1\\xCF").Visible=-1\r
Sheets("Dashboard \\xCF\\x27\\xD4\\xC8\\xE6\\xD1\\xCF").Select\r
Call OpenMenu\r
End Sub\r
Sub ShowReports()\r
On Error Resume Next\r
Sheets("Reports \\xAF\\xD2\\x27\\xD1\\xD4\\x87\\x27").Visible=-1\r
Sheets("Reports \\xAF\\xD2\\x27\\xD1\\xD4\\x87\\x27").Select\r
Call OpenMenu\r
End Sub\r
Sub ShowReturnForm()\r
On Error Resume Next\r
Sheets(15).Visible=-1\r
Sheets(15).Select\r
Call OpenMenu\r
End Sub\r
"""

def main():
    with zipfile.ZipFile(SAMPLE_XLSM, 'r') as z:
        vba_bin = z.read('xl/vbaProject.bin')
    tmp_orig="/tmp/vba_v6_orig.bin"
    tmp_new="/tmp/vba_v6_new.bin"
    with open(tmp_orig,'wb') as f:
        f.write(vba_bin)
    import shutil
    shutil.copy(tmp_orig, tmp_new)
    ole = olefile.OleFileIO(tmp_new, write_mode=True)

    def repl(path, src):
        try:
            old=ole.openstream(path).read()
            new_data=msovba.replace_module_source(old, src)
            ole.write_stream(path, new_data)
            print(f"OK {path}: {len(old)}->{len(new_data)} src={len(src)}")
            return True
        except Exception as e:
            print(f"FAIL {path}: {e}")
            return False

    repl('VBA/Module1', MODULE1_CODE)
    repl('VBA/ThisWorkbook', THISWORKBOOK_CODE)
    repl('VBA/Login', LOGIN_CODE)
    repl('VBA/Module2', MODULE2_CODE)
    repl('VBA/Module3', MODULE3_CODE)
    repl('VBA/Module4', MODULE4_CODE)
    repl('VBA/Module5', MODULE5_CODE)

    ole.close()
    with open(tmp_new,'rb') as f:
        new_vba=f.read()
    print(f"New vba size {len(new_vba)}")

    # Embed into xlsx
    with zipfile.ZipFile(BASE_XLSX,'r') as zin:
        files={n:zin.read(n) for n in zin.namelist()}
    ct=files['[Content_Types].xml'].decode()
    ct=ct.replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
                  'application/vnd.ms-excel.sheet.macroEnabled.main+xml')
    if 'vbaProject.bin' not in ct:
        ct=ct.replace('</Types>','<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/><Override PartName="/xl/vbaProjectSignature.bin" ContentType="application/vnd.ms-office.vbaProjectSignature"/></Types>')
    files['[Content_Types].xml']=ct.encode()
    files['xl/vbaProject.bin']=new_vba
    rels_path='xl/_rels/workbook.xml.rels'
    if rels_path in files:
        rels=files[rels_path].decode()
        if 'vbaProject' not in rels:
            rels=rels.replace('</Relationships>','<Relationship Id="rId999" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>')
            files[rels_path]=rels.encode()
    else:
        files[rels_path]=b'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId999" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>'

    with zipfile.ZipFile(OUT_XLSM,'w',zipfile.ZIP_DEFLATED) as zout:
        for name,data in files.items():
            zout.writestr(name,data)
    print(f"Wrote {OUT_XLSM} {os.path.getsize(OUT_XLSM)}")
    shutil.copy(OUT_XLSM, FINAL1)
    shutil.copy(OUT_XLSM, FINAL2)
    print(f"Copied to {FINAL1} and {FINAL2}")

if __name__=='__main__':
    main()
