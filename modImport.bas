Attribute VB_Name = "modImport"
Sub GetFileName()


Range("fileCount").Formula = 95
Range("L3").Formula = "=IFERROR(INDEX(FileList,ROW()-ROW(L$2)),""N/A"")"
Range("L3").AutoFill Destination:=Range("L3").Resize(Range("fileCount").Value)
Range("H5").Formula = "=COUNTA(ListofFiles)-COUNTIF(ListofFiles,""N/A"")"
Range("FileName").Formula = "=INDEX(ListOfFiles,H5)"
Range("LatestFileName").Formula = "=SUBSTITUTE(FileName,"".csv"","""")"
Range("LatestSheetName").Formula = "=LEFT(G4,LEN(G4)-5)"


'this is identifying the latest sheet in the file explorer
'ensure that the file path in cell G2 ("Import" Sheet) is correct if wanting to pull sheet through from elsewhere


End Sub

Sub OpenMostRecentFile()

Dim myFile As String
Dim myRecentFile As String
Dim myMostRecentFile As String
Dim recentData As Date

Dim myDirectory As String
myDirectory = "X:\Bet Tribe\Trading\FOOTBALL DEPARTMENT\RAB Reports\Monkeybot Reports\JIRA Data" 'change file link where appropriate as this opens files in specific location

Dim fileExtension As String
fileExtension = "*.csv"

Application.ScreenUpdating = False
Application.DisplayAlerts = False

If Right(myDirectory, 1) <> "\" Then myDirectory = myDirectory & "\"

myFile = Dir(myDirectory & fileExtension)

If myFile <> "" Then
    myRecentFile = myFile
    recentDate = FileDateTime(myDirectory & myFile)
Do While myFile <> ""
    If FileDateTime(myDirectory & myFile) > recentDate Then
    myRecentFile = myFile
    recentDate = FileDateTime(myDirectory & myFile)
End If
myFile = Dir
Loop
End If
myMostRecentFile = myRecentFile
Workbooks.Open fileName:=myDirectory & myMostRecentFile

Workbooks("Monkeybot Report Template.xlsm").Activate 'if the template is renamed, it will need to be updated here otherwise code will fail

End Sub

Sub CopyMethod()

Dim fileName As String: fileName = Range("LatestFileName")
Dim sheetName As String: sheetName = Range("LatestSheetName")

Workbooks(fileName & ".csv").Worksheets(sheetName).Cells.Copy _
Workbooks("Monkeybot Report Template").Worksheets("RawData").Range("A1")
Sheets("RawData").Visible = True
    Sheets("RawData").Select
    Cells.Select
    Selection.RowHeight = 15
Worksheets("Import").Activate
Workbooks(fileName & ".csv").Close
Sheets("RawData").Visible = False
MsgBox ("Import Complete")

End Sub


