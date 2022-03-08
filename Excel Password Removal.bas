Attribute VB_Name = "Module1"
Sub get_files()
    Dim ws As Worksheet
    Dim strFile As String
    Dim oShell: Set oShell = CreateObject("Shell.Application")
    Dim oDir As Variant
    Dim i As Long, x As Long
    Dim icount As Integer: icount = 1
    
    Set ws = Sheets("Panel")
    
    baseFolder = ws.Range("folder_path").Value & "\"
    Set oDir = oShell.Namespace(baseFolder)
    
    For i = 0 To 288
        If oDir.GetDetailsOf(oDir.Items, i) = "Program name" Then
            x = i
            Exit For
        End If
    Next i
    
    For Each sFile In oDir.Items
        If oDir.GetDetailsOf(sFile, x) = "" Then
            ws.Cells(icount + 3, 1).Value = icount
            ws.Cells(icount + 3, 2).Value = sFile.Name
            icount = icount + 1
        End If
    Next
    
    ws.Range(ws.Cells(4, 1), ws.Cells(icount + 2, 4)).Style = "table_cell"
    
End Sub

Sub decryptExcel()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim noentry As Boolean: noentry = False
    Dim irow As Integer: irow = 4
    
    Set ws = Sheets("Panel")
    baseFolder = ws.Range("folder_path").Value & "\"
    
    Application.DisplayAlerts = False
    Do While IsEmpty(ws.Cells(irow, 1).Value) = False
        On Error Resume Next
        filename = ws.Cells(irow, 2).Value
        pw = ws.Cells(irow, 3).Value
        Set wb = Workbooks.Open(baseFolder & filename, Password:=pw)
        If Err.Number = 0 Then
            wb.SaveAs baseFolder & filename, Password:=""
            wb.Close
            ws.Cells(irow, 4).Value = "Success"
            ws.Cells(irow, 4).Style = "success_decrypt"
        Else
            Err.Clear
            ws.Cells(irow, 4).Value = "Fail"
            ws.Cells(irow, 4).Style = "failed_decrypt"
        End If
        irow = irow + 1
    Loop
    Application.DisplayAlerts = True
End Sub
