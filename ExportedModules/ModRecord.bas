Attribute VB_Name = "ModRecord"
Sub UpdateGameRecord(score As Long, level As Long, rowsCleared As Integer, Quads As Integer)
    Dim ws_rd As Worksheet
    Set ws_rd = Worksheets("Record")
    Dim historyList As MSForms.ListBox

    Set historyList = ws_rd.OLEObjects("ListBox1").Object

    ' Create the history data
    Dim displayText As String
    displayText = Format(Now, "yyyy-mm-dd hh:mm:ss") & _
                   " - Score: " & score & _
                   ", Level: " & level & _
                   ", Rows Cleared: " & rowsCleared & _
                   ", Quads: " & Quads

    ' Adding Histroy to listBox
    historyList.AddItem displayText
    Debug.Print displayText
End Sub
Sub ClearHistory()
    Dim ws_rd As Worksheet
    Set ws_rd = Worksheets("Record")
    Dim historyList As MSForms.ListBox

    Set historyList = ws_rd.OLEObjects("ListBox1").Object

    historyList.Clear
End Sub

