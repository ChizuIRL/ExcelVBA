Attribute VB_Name = "Module2"
Sub copytocolend()
Dim carmanuf As String
Dim carmodel As String
Dim color As String

    carmanuf = Range("B2").Value
    carmodel = Range("C2").Value
    color = Range("D2").Value

    Range("H1").Select
    Selection.End(xlDown).Offset(1, 0).Select
    Selection.Value = carmanuf & " " & carmodel & ", " & color

End Sub
