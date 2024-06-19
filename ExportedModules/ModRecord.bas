Attribute VB_Name = "ModRecord"
Sub UpdateGameRecord(score As Long, level As Long, rowsCleared As Integer, Quads As Integer)
    Dim ws_rd As Worksheet
    Set ws_rd = Worksheets("Record")
    Dim historyList As MSForms.ListBox
    Set historyList = ws_rd.OLEObjects("ListBox1").Object

    Dim displayText As String
    displayText = userName & ", Score: " & score & _
                   ", Level: " & level & _
                   ", Rows: " & rowsCleared & _
                   ", Quads: " & Quads

    historyList.AddItem displayText
    Call SortListBoxByScore(historyList)
    Debug.Print displayText
End Sub

Sub SortListBoxByScore(historyList As MSForms.ListBox)
    Dim i As Integer, j As Integer
    Dim temp As String
    Dim scores() As Long
    Dim items() As String
    Dim n As Integer

    n = historyList.ListCount
    ReDim scores(1 To n)
    ReDim items(1 To n)

    For i = 1 To n
        items(i) = historyList.List(i - 1)
        items(i) = Trim(Replace(items(i), "1st: ", ""))
        items(i) = Trim(Replace(items(i), "2nd: ", ""))
        items(i) = Trim(Replace(items(i), "3rd: ", ""))
        items(i) = Trim(Replace(items(i), i & "th: ", ""))

        Dim scoreParts As Variant
        scoreParts = Split(items(i), "Score: ")
        If UBound(scoreParts) > 0 Then
            Dim scoreValueParts As Variant
            scoreValueParts = Split(scoreParts(1), ",")
            If UBound(scoreValueParts) >= 0 Then
                scores(i) = CLng(Trim(scoreValueParts(0)))
            Else
                scores(i) = 0
            End If
        Else
            scores(i) = 0
        End If
    Next i

    For i = 1 To n - 1
        For j = i + 1 To n
            If scores(i) < scores(j) Then
                temp = scores(i)
                scores(i) = scores(j)
                scores(j) = temp
                temp = items(i)
                items(i) = items(j)
                items(j) = temp
            End If
        Next j
    Next i

    historyList.Clear
    For i = 1 To n
        If i > 5 Then Exit For
        Dim rank As String
        Select Case i
            Case 1
                rank = "1st: "
            Case 2
                rank = "2nd: "
            Case 3
                rank = "3rd: "
            Case Else
                rank = i & "th: "
        End Select
        historyList.AddItem rank & items(i)
    Next i
End Sub


Sub ClearHistory()
    Dim ws_rd As Worksheet
    Set ws_rd = Worksheets("Record")
    Dim historyList As MSForms.ListBox

    Set historyList = ws_rd.OLEObjects("ListBox1").Object

    historyList.Clear
    Sheets("Data").Cells.Clear
End Sub

Sub SaveListBoxData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Data")  ' A hidden sheet for storing data
    Dim lb As MSForms.ListBox
    Set lb = Worksheets("Record").OLEObjects("ListBox1").Object
    Dim i As Integer
    
    ws.Cells.Clear  ' Clear existing data
    
    For i = 0 To lb.ListCount - 1
        ws.Cells(i + 1, 1).Value = lb.List(i)
    Next i
End Sub

Sub LoadListBoxData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Data")
    Dim lb As MSForms.ListBox
    Set lb = Worksheets("Record").OLEObjects("ListBox1").Object
    Dim i As Integer
    Dim lastRow As Integer

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row  ' Find the last row with data
    
    lb.Clear  ' Clear current ListBox contents before loading
    
    For i = 1 To lastRow
        lb.AddItem ws.Cells(i, 1).Value
    Next i
End Sub


