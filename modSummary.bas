Attribute VB_Name = "modSummary"
Sub generateTotals()

Application.ScreenUpdating = False

Worksheets("Summary").Activate

Range("RequestsReceived").Formula = "=totReq"
Range("requestsRejected").Formula = "=COUNTIF('Formatted Data'!B:B,""Rejected"")+COUNTIF('Formatted Data'!C:C,""Rejected"")"
Range("totalTime").Formula = "=(totHrs & ""hrs"")"
Range("avgTime").Formula = "=avgResp& "" mins"""

End Sub

Sub componentByDay()

Application.ScreenUpdating = False

Sheets("Lists").Select
Range("H4").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("Summary").Select
Range("B20").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Range("C20").Formula = "=COUNTIFS('Formatted Data'!$D:$D,""*""&Summary!$B20&""*"",'Formatted Data'!$H:$H,Summary!$C$19)"
Range("C20").AutoFill Destination:=Range("C20").Resize(Range("componentCount").Value)

Range("D20").Formula = "=COUNTIFS('Formatted Data'!$D:$D,""*""&Summary!$B20&""*"",'Formatted Data'!$H:$H,Summary!$D$19)"
Range("D20").AutoFill Destination:=Range("D20").Resize(Range("componentCount").Value)

Range("E20").Formula = "=COUNTIFS('Formatted Data'!$D:$D,""*""&Summary!$B20&""*"",'Formatted Data'!$H:$H,Summary!$E$19)"
Range("E20").AutoFill Destination:=Range("E20").Resize(Range("componentCount").Value)

Range("F20").Formula = "=COUNTIFS('Formatted Data'!$D:$D,""*""&Summary!$B20&""*"",'Formatted Data'!$H:$H,Summary!$F$19)"
Range("F20").AutoFill Destination:=Range("F20").Resize(Range("componentCount").Value)

Range("G20").Formula = "=COUNTIFS('Formatted Data'!$D:$D,""*""&Summary!$B20&""*"",'Formatted Data'!$H:$H,Summary!$G$19)"
Range("G20").AutoFill Destination:=Range("G20").Resize(Range("componentCount").Value)

Range("H20").Formula = "=COUNTIFS('Formatted Data'!$D:$D,""*""&Summary!$B20&""*"",'Formatted Data'!$H:$H,Summary!$H$19)"
Range("H20").AutoFill Destination:=Range("H20").Resize(Range("componentCount").Value)

Range("I20").Formula = "=COUNTIFS('Formatted Data'!$D:$D,""*""&Summary!$B20&""*"",'Formatted Data'!$H:$H,Summary!$I$19)"
Range("I20").AutoFill Destination:=Range("I20").Resize(Range("componentCount").Value)

Range("J20").Formula = "=SUM(C20:I20)"
Range("J20").AutoFill Destination:=Range("J20").Resize(Range("componentCount").Value)

Range("K20").Formula = "=J20/totReq"
Range("K20").AutoFill Destination:=Range("K20").Resize(Range("componentCount").Value)

Range("C48").Formula = "=SUM(C20:C43)"
Range("D48").Formula = "=SUM(D20:D43)"
Range("E48").Formula = "=SUM(E20:E43)"
Range("F48").Formula = "=SUM(F20:F43)"
Range("G48").Formula = "=SUM(G20:G43)"
Range("H48").Formula = "=SUM(H20:H43)"
Range("I48").Formula = "=SUM(I20:I43)"

Range("C49").Formula = "=SUMIFS('Formatted Data'!$G:$G,'Formatted Data'!$K:$K,""Y"",'Formatted Data'!$H:$H,Summary!C$19)"
Range("D49").Formula = "=SUMIFS('Formatted Data'!$G:$G,'Formatted Data'!$K:$K,""Y"",'Formatted Data'!$H:$H,Summary!D$19)"
Range("E49").Formula = "=SUMIFS('Formatted Data'!$G:$G,'Formatted Data'!$K:$K,""Y"",'Formatted Data'!$H:$H,Summary!E$19)"
Range("F49").Formula = "=SUMIFS('Formatted Data'!$G:$G,'Formatted Data'!$K:$K,""Y"",'Formatted Data'!$H:$H,Summary!F$19)"
Range("G49").Formula = "=SUMIFS('Formatted Data'!$G:$G,'Formatted Data'!$K:$K,""Y"",'Formatted Data'!$H:$H,Summary!G$19)"
Range("H49").Formula = "=SUMIFS('Formatted Data'!$G:$G,'Formatted Data'!$K:$K,""Y"",'Formatted Data'!$H:$H,Summary!H$19)"
Range("I49").Formula = "=SUMIFS('Formatted Data'!$G:$G,'Formatted Data'!$K:$K,""Y"",'Formatted Data'!$H:$H,Summary!I$19)"

Range("C50").Formula = "=C48/RequestsReceived"
Range("D50").Formula = "=D48/RequestsReceived"
Range("E50").Formula = "=E48/RequestsReceived"
Range("F50").Formula = "=F48/RequestsReceived"
Range("G50").Formula = "=G48/RequestsReceived"
Range("H50").Formula = "=H48/RequestsReceived"
Range("I50").Formula = "=I48/RequestsReceived"

End Sub

Sub GetPopularComponents()

Dim rng As Range, cell As Range
Dim firstVal As Double, secondVal As Double, thirdVal As Double, fourthVal As Double, fifthVal As Double

Worksheets("Summary").Activate

Set rng = [J20:J43]

firstVal = Application.WorksheetFunction.Large(rng, 1)
secondVal = Application.WorksheetFunction.Large(rng, 2)
thirdVal = Application.WorksheetFunction.Large(rng, 3)
fourthVal = Application.WorksheetFunction.Large(rng, 4)
fifthVal = Application.WorksheetFunction.Large(rng, 5)

Worksheets("Lists").Visible = True

Worksheets("Lists").Activate
Range("L4").Value = firstVal
Range("L5").Value = secondVal
Range("L6").Value = thirdVal
Range("L7").Value = fourthVal
Range("L8").Value = fifthVal

Range("K4").Formula = "=INDEX(Summary!$B$20:$B$43,MATCH(Lists!L4,Summary!$J$20:$J$43,0))"
Range("K4").AutoFill Destination:=Range("K4").Resize(numRows + 5)

Range("K3").Value = "Top Components"
Range("L3").Value = "Count"


End Sub

Sub RejectedBreakdown()

Worksheets("Summary").Activate

Range("M20").Formula = "=B20"
Range("M20").AutoFill Destination:=Range("M20").Resize(Range("componentCount").Value)

Range("N20").Formula = "=COUNTIF('Formatted Data'!D:D,Summary!$M$19&"" / ""&M20)+COUNTIF('Formatted Data'!D:D,Summary!M20& "" / ""&Summary!$M$19)"
Range("N20").AutoFill Destination:=Range("N20").Resize(Range("componentCount").Value)

Range("O20").Formula = "=N20/$B$7"
Range("O20").AutoFill Destination:=Range("O20").Resize(Range("componentCount").Value)


End Sub

Sub TraderSummary()

Worksheets("Summary").Activate

Range("B54").Value = "Trader"
Range("C54").Value = "Total"
Range("D54").Value = "%"
Range("E54").Value = "Time(hrs)"
Range("F54").Value = "AvgTime(mins)"

Range("B75").Value = "Trader"
Range("C75").Formula = "=Lists!K4"
Range("D75").Formula = "=Lists!K5"
Range("E75").Formula = "=Lists!K6"
Range("F75").Formula = "=Lists!K7"
Range("G75").Formula = "=Lists!K8"

Range("B95").Value = "Trader"
Range("C95").Formula = "=Lists!K4"
Range("D95").Formula = "=Lists!K5"
Range("E95").Formula = "=Lists!K6"
Range("F95").Formula = "=Lists!K7"
Range("G95").Formula = "=Lists!K8"

Sheets("Lists").Select
Range("E4").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("Summary").Select
Range("B55").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False


Range("C55").Formula = "=COUNTIF('Formatted Data'!$E:$E,Summary!B55)"
Range("C55").AutoFill Destination:=Range("C55").Resize(Range("TraderCount").Value)

Range("D55").Formula = "=C55/RequestsReceived"
Range("D55").AutoFill Destination:=Range("D55").Resize(Range("TraderCount").Value)

Range("E55").Formula = "=SUMIFS('Formatted Data'!$G:$G,'Formatted Data'!$K:$K,""Y"",'Formatted Data'!$E:$E,Summary!B55)/60"
Range("E55").AutoFill Destination:=Range("E55").Resize(Range("TraderCount").Value)

Range("F55").Formula = "=(E55*60)/C55"
Range("F55").AutoFill Destination:=Range("F55").Resize(Range("TraderCount").Value)

Sheets("Lists").Select
Range("E4").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("Summary").Select
Range("B76").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Range("C76").Formula = "=COUNTIFS('Formatted Data'!$E:$E,Summary!B76,'Formatted Data'!$D:$D,""*""&Summary!$C$75&""*"")"
Range("C76").AutoFill Destination:=Range("C76").Resize(Range("TraderCount").Value)

Range("D76").Formula = "=COUNTIFS('Formatted Data'!$E:$E,Summary!B76,'Formatted Data'!$D:$D,""*""&Summary!$D$75&""*"")"
Range("D76").AutoFill Destination:=Range("D76").Resize(Range("TraderCount").Value)

Range("E76").Formula = "=COUNTIFS('Formatted Data'!$E:$E,Summary!B76,'Formatted Data'!$D:$D,""*""&Summary!$E$75&""*"")"
Range("E76").AutoFill Destination:=Range("E76").Resize(Range("TraderCount").Value)

Range("F76").Formula = "=COUNTIFS('Formatted Data'!$E:$E,Summary!B76,'Formatted Data'!$D:$D,""*""&Summary!$F$75&""*"")"
Range("F76").AutoFill Destination:=Range("F76").Resize(Range("TraderCount").Value)

Range("G76").Formula = "=COUNTIFS('Formatted Data'!$E:$E,Summary!B76,'Formatted Data'!$D:$D,""*""&Summary!$G$75&""*"")"
Range("G76").AutoFill Destination:=Range("G76").Resize(Range("TraderCount").Value)

Sheets("Lists").Select
Range("E4").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("Summary").Select
Range("B96").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Range("C96").Formula = "=IFERROR(ROUND(AVERAGEIFS('Formatted Data'!$G:$G,'Formatted Data'!$E:$E,Summary!B96,'Formatted Data'!$D:$D,""*""&Summary!$C$95&""*"",'Formatted Data'!$K:$K,""Y""),2),"""")"
Range("C96").AutoFill Destination:=Range("C96").Resize(Range("TraderCount").Value)

Range("D96").Formula = "=IFERROR(ROUND(AVERAGEIFS('Formatted Data'!$G:$G,'Formatted Data'!$E:$E,Summary!B96,'Formatted Data'!$D:$D,""*""&Summary!$D$95&""*"",'Formatted Data'!$K:$K,""Y""),2),"""")"
Range("D96").AutoFill Destination:=Range("D96").Resize(Range("TraderCount").Value)

Range("E96").Formula = "=IFERROR(ROUND(AVERAGEIFS('Formatted Data'!$G:$G,'Formatted Data'!$E:$E,Summary!B96,'Formatted Data'!$D:$D,""*""&Summary!$E$95&""*"",'Formatted Data'!$K:$K,""Y""),2),"""")"
Range("E96").AutoFill Destination:=Range("E96").Resize(Range("TraderCount").Value)

Range("F96").Formula = "=IFERROR(ROUND(AVERAGEIFS('Formatted Data'!$G:$G,'Formatted Data'!$E:$E,Summary!B96,'Formatted Data'!$D:$D,""*""&Summary!$F$95&""*"",'Formatted Data'!$K:$K,""Y""),2),"""")"
Range("F96").AutoFill Destination:=Range("F96").Resize(Range("TraderCount").Value)

Range("G96").Formula = "=IFERROR(ROUND(AVERAGEIFS('Formatted Data'!$G:$G,'Formatted Data'!$E:$E,Summary!B96,'Formatted Data'!$D:$D,""*""&Summary!$G$95&""*"",'Formatted Data'!$K:$K,""Y""),2),"""")"
Range("G96").AutoFill Destination:=Range("G96").Resize(Range("TraderCount").Value)

Range("Q3").Formula = "=C72"
Range("Q5").Formula = "=D72"
Range("Q7").Formula = "=E72"
Range("Q9").Formula = "=F72"
Range("Q11").Formula = "=G72"

End Sub

Sub FormatTables()

' Trader Requests By Component Formatting

    Range("B74:G92").Select
    
    With Selection.Borders(xlEdgeLeft)
    .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeTop)
    .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeBottom)
    .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeRight)
    .Weight = xlMedium
    End With
    
    Range("B75:G92").Select
    
    With Selection.Borders(xlEdgeLeft)
    .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeTop)
    .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeBottom)
    .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeRight)
    .Weight = xlMedium
    End With
    
    ' Trader Response Times By Component Formatting
    
    Range("B94:G111").Select
    
    With Selection.Borders(xlEdgeLeft)
    .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeTop)
    .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeBottom)
    .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeRight)
    .Weight = xlMedium
    End With
    
    Range("B95:G111").Select
    
    With Selection.Borders(xlEdgeLeft)
    .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeTop)
    .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeBottom)
    .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeRight)
    .Weight = xlMedium
    End With


    
End Sub


