Attribute VB_Name = "ModExtraFeatures"
Public IsGamePaused As Boolean
Public PausedFlag As Boolean
Public FeatureLimit As Integer
Public UsedTime As Integer
Public ButtonFlag As Integer

Sub PauseGame()
    ' Assume TimID is your timer identifier
    Call KillTimer(0&, TimID)
    IsGamePaused = True
End Sub
Sub PauseAndRestart()
    If PausedFlag = False Then
        Call KillTimer(0&, TimID)
        PausedFlag = True
    Else
        TimID = SetTimer(0&, 0&, MilSec, AddressOf TimerProcedure)
        PausedFlag = False
    End If
End Sub
Sub ResumeGame()
    ' Reset the timer with the same interval
    TimID = SetTimer(0&, 0&, MilSec, AddressOf TimerProcedure)
    IsGamePaused = False
    Call AddBlock(CurBlo.X, CurBlo.Y, 1)
End Sub
Sub RegenerateNextBlockAndContinue()
    If Not IsGamePaused Then
        Call PauseGame
    End If
    Debug.Print UsedTime
    Debug.Print FeatureLimit
    If UsedTime < FeatureLimit Then
        UsedTime = UsedTime + 1
    Else
        MsgBox "Usage limit exceeded."
        Exit Sub
    End If
    
    ' Generate a new next block
    Call RemoveBlock(CurBlo.X, CurBlo.Y)
    Call GenerateBlocks(1)
    

    ' Resume the game immediately after generating the new block
    ResumeGame

End Sub
Sub UpdateGameRecord(score As Long, level As Long, rowsCleared As Integer, Quads As Integer)
    Dim ws_rd As Worksheet
    Set ws_rd = Worksheets("Game Records")
    Dim historyList As MSForms.ListBox

    Set historyList = ws_rd.OLEObjects("ListBox1").Object

    ' 建立要顯示的訊息
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
    Set ws_rd = Worksheets("Game Records")
    Dim historyList As MSForms.ListBox

    Set historyList = ws_rd.OLEObjects("ListBox1").Object

    historyList.Clear
End Sub

