Attribute VB_Name = "Module5"
Sub CopyalltoMasterdata()
Dim i As Integer
Dim o As Integer
Dim usdrate As String
Dim sheettitle As String

usdrate = InputBox("Input the USD Value", "USD Rate", "For Example: 50.25")
Sheets.Add(After:=Sheets("December")).Name = "Annual Master Data"
For i = 3 To Worksheets.Count - 1

Worksheets(i).Select
sheettitle = ActiveSheet.Name
Worksheets("Annual Master Data").Select

Range("A1500").Select
Selection.End(xlUp).Select
ActiveCell.Offset(3, 0).Select
ActiveCell.Value = sheettitle
Selection.Font.Bold = True

Worksheets(i).Select
Range("A4").Select
Selection.CurrentRegion.Copy
Sheets("Annual Master Data").Select
ActiveCell.Offset(2, 0).Select
ActiveSheet.Paste

Next i
On Error Resume Next


  Columns("A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
  Cells.Select
  Selection.Font.Name = "Noto Sans JP"
  Columns("M").EntireColumn.Select
  Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
  Cells.Select
  ActiveSheet.Range("$A$1:$N$5640").RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14), Header:=xlNo
  Rows(2).Find("Car Cost").Offset(0, 1).EntireColumn.Insert
  Columns("N").EntireColumn.Select
  Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
  Selection.NumberFormat = "_-[$$-en-US]* #,##0.00_ ;_-[$$-en-US]* -#,##0.00 ;_-[$$-en-US]* ""-""??_ ;_-@_ "
  Range("A1").Select
  Selection.EntireRow.Insert
  Selection.EntireRow.Insert
  Selection.EntireRow.Insert
  
  

o = 5
Do While Cells(o, 1).Value <> ""
  Cells(o, 14).Value = Cells(o, 13).Value / usdrate
  o = o + 1
Loop

Range("N5").Value = "Car Cost USD"
Range("N5").Select.Font.Bold = True
Rows(3).Find("Car Cost USD").EntireColumn.Select
Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
Selection.NumberFormat = "_-[$$-en-US]* #,##0.00_ ;_-[$$-en-US]* -#,##0.00 ;_-[$$-en-US]* ""-""??_ ;_-@_ "
Rows("1:4").Select
Selection.NumberFormat = "General"
Cells.Select
Selection.Columns.AutoFit
Range("A1").Select

End Sub
