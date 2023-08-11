Attribute VB_Name = "Module7"
Public Sub Convertphptousd()
Dim i As Integer
Rows(3).Find("Car Cost").Offset(0, 1).EntireColumn.Insert

i = 4
Do While Cells(i, 1).Value <> ""
Cells(i, 7).Value = Cells(i, 6).Value / Cells(2, 3)
i = i + 1
Loop

    Range("G3").Value = "Car Cost USD"
    Range("G3").Select
    Selection.Font.Bold = True
    Rows(3).Find("Car Cost USD").EntireColumn.Select
    Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Selection.NumberFormat = "_-[$$-en-US]* #,##0.00_ ;_-[$$-en-US]* -#,##0.00 ;_-[$$-en-US]* ""-""??_ ;_-@_ "
    Rows("1:3").Select
    Selection.NumberFormat = "General"
    Columns("G:G").Select
    Selection.Columns.AutoFit
    Range("A1").Select
End Sub
