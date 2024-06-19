Attribute VB_Name = "ModExtraFeatures"
Public IsGamePaused As Boolean
Public PausedFlag As Boolean
Public FeatureLimit As Integer
Public UsedTime As Integer

Sub TogglePauseResumeGame()
    Call ClickSoundEffect
    If IsGamePaused Then
        IsGamePaused = False
        Call DrawPlayingField(1)
        Call ResumeGameTimer
    Else
        IsGamePaused = True
        Call PauseGameTimer
    End If
End Sub

Sub PauseGameTimer()
    Call KillTimer(0&, TimID)
    IsGamePaused = True
End Sub

Sub ResumeGameTimer()
    TimID = SetTimer(0&, 0&, MilSec, AddressOf TimerProcedure)
    IsGamePaused = False
End Sub

Sub RegenerateNextBlockAndContinue()
    Call ClickSoundEffect
    If Not IsGamePaused Then
        Call PauseGameTimer
    End If

    If UsedTime < FeatureLimit Then
        UsedTime = UsedTime + 1
        Call RemoveBlock(CurBlo.X, CurBlo.Y)
        Call GenerateBlocks(1)
        Call ResumeGameTimer
    Else
        MsgBox "Too many times, bro!"
        Call ResumeGameTimer
    End If
End Sub

