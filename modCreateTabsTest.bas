Attribute VB_Name = "modCreateTabsTest"
Sub CreateNewTabs() 'Not currently ready for release

Dim Sheet1 As Worksheet
Dim Sheet2 As Worksheet
Dim Sheet3 As Worksheet
Dim Sheet4 As Worksheet
Dim Sheet5 As Worksheet

Cmp1 = Worksheets("Summary").Range("C72")
Cmp2 = Worksheets("Summary").Range("D72")
Cmp3 = Worksheets("Summary").Range("E72")
Cmp4 = Worksheets("Summary").Range("F72")
Cmp5 = Worksheets("Summary").Range("G72")


Set Sheet1 = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
Sheet1.Name = Cmp1

Set Sheet2 = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
Sheet2.Name = Cmp2

Set Sheet3 = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
Sheet3.Name = Cmp3

Set Sheet4 = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
Sheet4.Name = Cmp4

Set Sheet5 = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
Sheet5.Name = Cmp5

End Sub

Sub PopulateSheet1() 'not currently in use

Crit1 = Worksheets("Lists").Range("K4")
Sht1 = Worksheets("Lists").Range("K4")

Worksheets(Sht1).Columns("A:C").ColumnWidth = 25
Worksheets(Sht1).Columns("D:D").ColumnWidth = 250

Worksheets("Formatted Data").Activate
Worksheets("Formatted Data").Range("A:L").Select
Selection.AutoFilter Field:=4, Criteria1:=(Crit1) & "*"
Range("A1, D1, H1, L1").EntireColumn.Select
Selection.Copy
Worksheets(Sht1).Activate
Range("A1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False
Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
SkipBlanks:=False, Transpose:=False


End Sub


Sub PopulateSheet2() 'not currently in use

Crit2 = Worksheets("Lists").Range("K5")
Sht2 = Worksheets("Lists").Range("K5")

Worksheets(Sht2).Columns("A:C").ColumnWidth = 25
Worksheets(Sht2).Columns("D:D").ColumnWidth = 250

Worksheets("Formatted Data").Activate
Worksheets("Formatted Data").Range("A:L").Select
Selection.AutoFilter Field:=4, Criteria1:=(Crit2) & "*"
Range("A1, D1, H1, L1").EntireColumn.Select
Selection.Copy
Worksheets(Sht2).Activate
Range("A1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False
Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
SkipBlanks:=False, Transpose:=False


End Sub


Sub PopulateSheet3() 'not currently in use

Crit3 = Worksheets("Lists").Range("K6")
Sht3 = Worksheets("Lists").Range("K6")

Worksheets(Sht3).Columns("A:C").ColumnWidth = 25
Worksheets(Sht3).Columns("D:D").ColumnWidth = 250

Worksheets("Formatted Data").Activate
Worksheets("Formatted Data").Range("A:L").Select
Selection.AutoFilter Field:=4, Criteria1:=(Crit3) & "*"
Range("A1, D1, H1, L1").EntireColumn.Select
Selection.Copy
Worksheets(Sht3).Activate
Range("A1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False
Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
SkipBlanks:=False, Transpose:=False
    
End Sub


Sub PopulateSheet4() 'not currently in use

Crit4 = Worksheets("Lists").Range("K7")
Sht4 = Worksheets("Lists").Range("K7")

Worksheets(Sht4).Columns("A:C").ColumnWidth = 25
Worksheets(Sht4).Columns("D:D").ColumnWidth = 250

Worksheets("Formatted Data").Activate
Worksheets("Formatted Data").Range("A:L").Select
Selection.AutoFilter Field:=4, Criteria1:=(Crit4) & "*"
Range("A1, D1, H1, L1").EntireColumn.Select
Selection.Copy
Worksheets(Sht4).Activate
Range("A1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False
Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
SkipBlanks:=False, Transpose:=False
    
End Sub


Sub PopulateSheet5() 'not currently in use

Crit5 = Worksheets("Lists").Range("K8")
Sht5 = Worksheets("Lists").Range("K8")

Worksheets(Sht5).Columns("A:C").ColumnWidth = 25
Worksheets(Sht5).Columns("D:D").ColumnWidth = 250

Worksheets("Formatted Data").Activate
Worksheets("Formatted Data").Range("A:L").Select
Selection.AutoFilter Field:=4, Criteria1:=(Crit5) & "*"
Range("A1, D1, H1, L1").EntireColumn.Select
Selection.Copy
Worksheets(Sht5).Activate
Range("A1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False
Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
SkipBlanks:=False, Transpose:=False

Worksheets("Formatted Data").AutoFilterMode = False

    
End Sub

Sub HideUnusedSheets()

Dim time As Double: time = Now()

Worksheets("Formatted Data").Visible = False
Worksheets("Lists").Visible = False
MsgBox "Setup Complete" & vbNewLine & "Time Taken: " & Now() - time
Worksheets("Summary").Activate

End Sub
