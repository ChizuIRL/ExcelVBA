Attribute VB_Name = "Module2"
Sub SortReg_Rep()
Attribute SortReg_Rep.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SortReg_Rep Macro
'

'
    Range("B2").Select
    ActiveWorkbook.Worksheets("SalesOrders").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SalesOrders").Sort.SortFields.Add2 Key:=Range( _
        "B2:B44"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("SalesOrders").Sort.SortFields.Add2 Key:=Range( _
        "C2:C44"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("SalesOrders").Sort
        .SetRange Range("A1:I44")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub Sort_Item()
Attribute Sort_Item.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SortbyItem Macro
'

'
    ActiveWorkbook.Worksheets("SalesOrders").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SalesOrders").Sort.SortFields.Add2 Key:=Range( _
        "D2:D44"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("SalesOrders").Sort
        .SetRange Range("A1:I44")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub Sort_mostUnits()
Attribute Sort_mostUnits.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SortbyUnits Macro
'

'
    ActiveWorkbook.Worksheets("SalesOrders").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SalesOrders").Sort.SortFields.Add2 Key:=Range( _
        "E2:E44"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("SalesOrders").Sort
        .SetRange Range("A1:I44")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub Sort_mostSubtotal()
Attribute Sort_mostSubtotal.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SortbySubtotal Macro
'

'
    ActiveWorkbook.Worksheets("SalesOrders").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SalesOrders").Sort.SortFields.Add2 Key:=Range( _
        "I2:I44"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("SalesOrders").Sort
        .SetRange Range("A1:I44")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub Sort_leastUnits()
Attribute Sort_leastUnits.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Sort_leastUnits Macro
'

'
    ActiveWorkbook.Worksheets("SalesOrders").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SalesOrders").Sort.SortFields.Add2 Key:=Range( _
        "E2:E44"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("SalesOrders").Sort
        .SetRange Range("A1:I44")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub Sort_leastSubtotal()
Attribute Sort_leastSubtotal.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Sort_leastSubtotal Macro
'

'
    ActiveWorkbook.Worksheets("SalesOrders").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SalesOrders").Sort.SortFields.Add2 Key:=Range( _
        "I2:I44"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("SalesOrders").Sort
        .SetRange Range("A1:I44")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub Sort_Date()
Attribute Sort_Date.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Sort_Date Macro
'

'
    ActiveWorkbook.Worksheets("SalesOrders").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SalesOrders").Sort.SortFields.Add2 Key:=Range( _
        "A2:A44"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("SalesOrders").Sort
        .SetRange Range("A1:I44")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub FIlter_Clear()
Attribute FIlter_Clear.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FIlter_Clear Macro
'

'
    ActiveSheet.ShowAllData
    Selection.AutoFilter
End Sub
Sub Filter_Q4()
Attribute Filter_Q4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Filter_Q4 Macro
'

'
    ActiveSheet.Range("$A$3:$I$46").AutoFilter Field:=1, Criteria1:=20, _
        Operator:=11, Criteria2:=0, SubField:=0
    ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort.SortFields.Add2 Key _
        :=Range("A3:A46"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A3").Select
    ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort.SortFields.Add2 Key _
        :=Range("A3:A46"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub Filter_Q1()
Attribute Filter_Q1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Filter_Q1 Macro
'

'
    Selection.AutoFilter
    ActiveSheet.Range("$A$3:$I$44").AutoFilter Field:=1, Criteria1:=17, _
        Operator:=11, Criteria2:=0, SubField:=0
    ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort.SortFields.Add2 Key _
        :=Range("A3:A44"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub Filter_Q2()
Attribute Filter_Q2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Filter_Q2 Macro
'

'
    Selection.AutoFilter
    ActiveSheet.Range("$A$3:$I$44").AutoFilter Field:=1, Criteria1:=18, _
        Operator:=11, Criteria2:=0, SubField:=0
    ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort.SortFields.Add2 Key _
        :=Range("A3:A44"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub Filter_Q3()
Attribute Filter_Q3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Filter_Q3 Macro
'

'
    Selection.AutoFilter
    ActiveSheet.Range("$A$3:$I$44").AutoFilter Field:=1, Criteria1:=19, _
        Operator:=11, Criteria2:=0, SubField:=0
    ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort.SortFields.Add2 Key _
        :=Range("A3:A44"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort.SortFields.Add2 Key _
        :=Range("A3:A44"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub Filter_withDiscounts()
Attribute Filter_withDiscounts.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Filter_withDiscounts Macro
'

'
    Range("A4:A5").Select
    Range("A5").Activate
    Selection.AutoFilter
    Selection.AutoFilter
    Range("A3").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$3:$I$44").AutoFilter Field:=8, Criteria1:=">50", _
        Operator:=xlAnd
End Sub
Sub Filter_SubtotalGr1000()
Attribute Filter_SubtotalGr1000.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Filter_SubtotalGr1000 Macro
'

'
    Range("A3").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$3:$I$44").AutoFilter Field:=9, Criteria1:=">=1000", _
        Operator:=xlAnd
End Sub
Sub Filter_UnitsGr50()
Attribute Filter_UnitsGr50.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Filter_UnitsGr50 Macro
'

'
    Range("A3").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$3:$I$44").AutoFilter Field:=5, Criteria1:=">=50", _
        Operator:=xlAnd
End Sub
Sub Filter_East()
Attribute Filter_East.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Filter_East Macro
'

'
    Range("A3").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$3:$I$44").AutoFilter Field:=2, Criteria1:="East"
    ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort.SortFields.Add2 Key _
        :=Range("A3:A44"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub Filter_Central()
Attribute Filter_Central.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Filter_North Macro
'

'
    Range("A3").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$3:$I$44").AutoFilter Field:=2, Criteria1:="Central"
    ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort.SortFields.Add2 Key _
        :=Range("A3:A44"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub Filter_West()
Attribute Filter_West.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Filter_West Macro
'

'
    Range("A3").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$3:$I$44").AutoFilter Field:=2, Criteria1:="West"
    ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort.SortFields.Add2 Key _
        :=Range("A3:A44"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort.SortFields.Add2 Key _
        :=Range("A3:A44"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub Filter_South()
Attribute Filter_South.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Filter_South Macro
'

'
    Range("A3").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$3:$I$44").AutoFilter Field:=2, Criteria1:="South"
    ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort.SortFields.Add2 Key _
        :=Range("A3:A44"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort.SortFields.Add2 Key _
        :=Range("A3:A44"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort.SortFields.Add2 Key _
        :=Range("A3:A44"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("SalesOrders").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
