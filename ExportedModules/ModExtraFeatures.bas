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
        ResumeGame
        Exit Sub
    End If
    
    ' Generate a new next block
    Call RemoveBlock(CurBlo.X, CurBlo.Y)
    Call GenerateBlocks(1)
    

    ' Resume the game immediately after generating the new block
    ResumeGame

End Sub


