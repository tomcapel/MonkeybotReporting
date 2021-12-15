Attribute VB_Name = "modTestFrequentWords"
Sub GenerateFrequentWords() 'this code is a TEST and not currently released

'word frequency
'Put the data in col A, run the code, the result is in col D:E.

Dim regEx As Object, matches As Object, x As Object, d As Object
Dim obj As New DataObject
Dim tx As String, z As String
Dim t, q, va
Dim i As Long

t = Timer
Range("A1", Cells(Rows.Count, "A").End(xlUp)).Copy
obj.GetFromClipboard
tx = obj.GetText
Application.CutCopyMode = False
tx = Replace(tx, "'", "___")
    
        Set regEx = CreateObject("VBScript.RegExp")
        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = True
            .Pattern = "\w+"
        End With

    Set d = CreateObject("scripting.dictionary")
    d.CompareMode = vbTextCompare
        
            Set matches = regEx.Execute(tx)
            
            For Each x In matches
                d(CStr(x)) = d(CStr(x)) + 1
            Next
                
If d.Count = 0 Then MsgBox "Nothing found": Exit Sub

'put the result in col D:E
Range("D:E").ClearContents
With Range("D2").Resize(d.Count, 2)
    If d.Count < 65536 Then 'Transpose function has a limit of 65536 item to process
        
        .Value = Application.Transpose(Array(d.Keys, d.items))
        
    Else
        
        ReDim va(1 To d.Count, 1 To 2)
        i = 0
            For Each q In d.Keys
                i = i + 1
                va(i, 1) = q: va(i, 2) = d(q)
            Next
        .Value = va
        
    End If
    .Replace What:="___", Replacement:="'", LookAt:=xlPart, SearchFormat:=False, ReplaceFormat:=False
    .Sort Key1:=.Cells(1, 2), Order1:=xlDescending, Key2:=.Cells(1, 1), Order2:=xlAscending, Header:=xlNo
    
End With

Range("D1") = "WORD"
Range("E1") = "FREQUENCY"
Range("D:E").Columns.AutoFit

Debug.Print Timer - t

End Sub
