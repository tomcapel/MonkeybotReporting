Attribute VB_Name = "modSetup"
Sub FormattedData() 'this sub covers initial setup, naming cells and adding key formulas to the 'formatted data' tab - config changes can be made here

Application.ScreenUpdating = False

Worksheets("Lists").Visible = True
Worksheets("Formatted Data").Visible = True
Worksheets("Summary").Visible = True


Worksheets("Summary").Activate
Range("totReq").Formula = "=COUNTA(RawData!A:A)-1"

Worksheets("Formatted Data").Activate

Range("A1").Value = "dateCreated"
Range("B1").Value = "requestComponent(1)"
Range("C1").Value = "requestComponent(2)"
Range("D1").Value = "componentString"
Range("E1").Value = "assignedTrader"
Range("F1").Value = "dateResolved"
Range("G1").Value = "resolveTime"
Range("H1").Value = "weekDay"
Range("I1").Value = "requestTime"
Range("J1").Value = "Time(Rnd)"
Range("K1").Value = "Include?"
Range("A2").Formula = "=RawData!Q2"
Range("A2").AutoFill Destination:=Range("A2").Resize(Range("totReq").Value)

Range("B2").Formula = "=IF(RawData!U2="""",""Not Assigned"",RawData!U2)"
Range("B2").AutoFill Destination:=Range("B2").Resize(Range("totReq").Value)

Range("C2").Formula = "=IF(RawData!V2="""","""",RawData!V2)"
Range("C2").AutoFill Destination:=Range("C2").Resize(Range("totReq").Value)

Range("D2").Formula = "=IF(C2="""",B2,B2&"" / ""&C2)"
Range("D2").AutoFill Destination:=Range("D2").Resize(Range("totReq").Value)

Range("E2").Formula = "=IFERROR(INDEX(TraderNames,MATCH(RawData!N2,TraderUsernames,0)),""Not Assigned"")"
Range("E2").AutoFill Destination:=Range("E2").Resize(Range("totReq").Value)

Range("F2").Formula = "=IF(OR(RawData!T2="""",RawData!T2=""Open Ticket""),""Open"",RawData!T2)"
Range("F2").AutoFill Destination:=Range("F2").Resize(Range("totReq").Value)

Range("G2").Formula = "=IF(F2=""Open"",""Open"",(F2-A2)*1440)"
Range("G2").AutoFill Destination:=Range("G2").Resize(Range("totReq").Value)

Range("H2").Formula = "=TEXT(A2,""DDDD"")"
Range("H2").AutoFill Destination:=Range("H2").Resize(Range("totReq").Value)

Range("I2").Formula = "=TEXT(A2,""HH:MM"")"
Range("I2").AutoFill Destination:=Range("I2").Resize(Range("totReq").Value)

Range("J2").Formula = "=MROUND(I2,1/24)"
Range("J2").AutoFill Destination:=Range("J2").Resize(Range("totReq").Value)

Range("K2").Formula = "=IF(OR(G2>60,G2<0),""N"",""Y"")"
Range("K2").AutoFill Destination:=Range("K2").Resize(Range("totReq").Value)

Range("L2").Formula = "=LEFT(RawData!X2,FIND(""From Slack"",RawData!X2)-1)"
Range("L2").AutoFill Destination:=Range("L2").Resize(Range("totReq").Value)

Worksheets("Lists").Activate
Range("B22").Value = "TotHrs"
Range("B23").Value = "AvgResp"
Range("B24").Value = "TotReq"
Range("C22").Formula = "=ROUND((SUMIF('Formatted Data'!K:K,""Y"",'Formatted Data'!G:G)/60),0)"
Range("C23").Formula = "=ROUND(AVERAGEIF('Formatted Data'!K:K,""Y"",'Formatted Data'!G:G),0)"
Range("B25").Value = "Earliest Date"
Range("B26").Value = "Latest Date"
Range("C25").Formula = "=MIN('Formatted Data'!A:A)"
Range("C26").Formula = "=MAX('Formatted Data'!A:A)"

Sheets("Formatted Data").Activate
Range("E2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("Lists").Select
Range("E4").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Range("E3").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.RemoveDuplicates Columns:=1, Header:=x1Yes

Range("F4").Formula = "=COUNTA(E:E)-1"

Range("E3").Value = "Trader"
Range("F3").Value = "Count"

Sheets("Formatted Data").Activate
Range("B2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("Lists").Select
Range("H4").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Range("H3").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.RemoveDuplicates Columns:=1, Header:=x1Yes

Range("I4").Formula = "=COUNTA(H:H)-1"

Range("H3").Value = "Component"
Range("I3").Value = "Count"

End Sub



