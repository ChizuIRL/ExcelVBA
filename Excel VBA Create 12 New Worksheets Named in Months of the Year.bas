Attribute VB_Name = "Module6"
Sub CreateMonthlyWorksheets()
    Dim months() As Variant
    Dim i As Integer
    
    months = Array("January", "February", "March", "April", "May", "June", _
                   "July", "August", "September", "October", "November", "December")
    
    For i = 0 To 11
        Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = months(i)
    Next i
    
    MsgBox "12 new worksheets have been created!", vbInformation
End Sub

