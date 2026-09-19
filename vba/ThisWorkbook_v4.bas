Attribute VB_Name = "ThisWorkbook"
' Place this code in ThisWorkbook module (double-click ThisWorkbook in VBA Project)
' It ensures only LOGIN is visible on open

Private Sub Workbook_Open()
    On Error Resume Next
    Application.ScreenUpdating = False
    ' Call login module handler
    Call QC_Full_v4.Workbook_Open_Handler
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    ' Optional: hide sheets again on close so file always opens to LOGIN
    Call QC_Full_v4.HideAllExceptLogin
    ThisWorkbook.Save
    On Error GoTo 0
End Sub
