Attribute VB_Name = "ModKeyHandlers"
Sub AssignKeys()
    On Error Resume Next
    With Application
        .OnKey "{DOWN}", "KeyDrop"
        .OnKey "{LEFT}", "KeyLeft"
        .OnKey "{RIGHT}", "KeyRight"
        .OnKey "{UP}", "KeyRotateCW"
        .OnKey " ", "KeyDrop"
        .OnKey "m", "RegenerateNextBlockAndContinue"
        .OnKey "s", "KeyDrop"
        .OnKey "a", "KeyLeft"
        .OnKey "d", "KeyRight"
        .OnKey "w", "KeyRotateCW"
        .OnKey "x", "KeyRotateCW"
        .OnKey "c", "KeyRotateCCW"
        .OnKey "p", "PauseAndRestart"
    End With
    On Error GoTo 0
End Sub

Sub RemoveKeyAssignations()
    On Error Resume Next
    With Application
        .OnKey "{DOWN}", ""
        .OnKey "{LEFT}", ""
        .OnKey "{RIGHT}", ""
        .OnKey "{UP}", ""
        .OnKey " ", ""
        .OnKey "m", ""
        .OnKey "s", ""
        .OnKey "a", ""
        .OnKey "d", ""
        .OnKey "w", ""
        .OnKey "x", ""
        .OnKey "c", ""
        .OnKey "p", ""
        
    End With
    On Error GoTo 0
End Sub

Sub KeyDrop()
    If GamSta = 1 Then
        Call MoveBlockDown(1)
    End If
End Sub

Sub KeyLeft()
    If GamSta = 1 Then
        Call MoveBlockRightLeft(-1)
    End If
End Sub

Sub KeyRight()
    If GamSta = 1 Then
        Call MoveBlockRightLeft(1)
    End If
End Sub

Sub KeyRotateCW()
    If GamSta = 1 Then
        Call RotateBlock(1)
    End If
End Sub

Sub KeyRotateCCW()
    If GamSta = 1 Then
        Call RotateBlock(-1)
    End If
End Sub
