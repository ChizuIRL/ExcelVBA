Attribute VB_Name = "Module1"
Sub Callallthreerows()

resp = MsgBox("Are you sure that this is the final template? Ie. Font, Bold/Italics", vbYesNo + vbQuestion)

If resp = vbYes Then
    Call CreateMonthlyWorksheets
    Call Copyformat
    Call Formatall
Else: resp = vbNo
On Error GoTo 0
End If




End Sub

Sub CreateMonthlyWorksheets()
    Dim months() As Variant
    Dim i As Integer
    
    months = Array("January", "February", "March", "April", "May", "June", _
                   "July", "August", "September", "October", "November", "December")
    
    For i = 0 To 11
    On Error Resume Next
        Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = months(i)
    
    Next i
    
End Sub

Sub Copyformat()

    Sheets("Sheet1").Activate
    Rows("1:3").Select
    Sheets(Array("Sheet1", "January", "February", "March", "April", "May", "June", _
                "July", "August", "September", "October", "November", "December")).Select
    Sheets("Sheet1").Activate
    ActiveWindow.SelectedSheets.FillAcrossSheets Range:=Selection, Type:=xlAll
    
    
End Sub
Sub Formatall()
Dim i As Integer
Dim shtitle As String

    For i = 2 To Worksheets.Count
    Worksheets(i).Select
    shtitle = ActiveSheet.Name
    
    Rows("1:3").Select
    Selection.NumberFormat = "General"
    Range("B2").Select
    Selection.Value = shtitle
    Range("A4").Select
    ActiveWindow.FreezePanes = True
    Range("A1:C1").Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .MergeCells = True
    End With
    Range("A2").Select
    Selection.HorizontalAlignment = xlLeft
    Selection.VerticalAlignment = xlBottom
    Cells.Select
    Selection.Columns.AutoFit
    Selection.Rows.AutoFit
    Range("A1").Select
    Next i
    
    Application.DisplayAlerts = False
    Sheets("Sheet1").Delete
    Application.DisplayAlerts = True
    Sheets("January").Select
    Range("A1").Select
    
End Sub
Sub vbboxed()

End Sub
