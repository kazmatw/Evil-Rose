Attribute VB_Name = "ModKeyHandlers2p"

Sub AssignKeys_2p()
    On Error Resume Next
    With Application
        'first player
        .OnKey "{DOWN}", "KeyDrop"
        .OnKey "{LEFT}", "KeyLeft"
        .OnKey "{RIGHT}", "KeyRight"
        .OnKey "{UP}", "KeyRotateCW"
        '.OnKey " ", "KeyDrop"
        .OnKey "m", "RegenerateNextBlockAndContinue"
        
        'second player
        .OnKey "s", "KeyDrop_2p"
        .OnKey "a", "KeyLeft_2p"
        .OnKey "d", "KeyRight_2p"
        .OnKey "w", "KeyRotateCW_2p"
        .OnKey "{TAB}", "RegenerateNextBlockAndContinue_2p"
        '.OnKey "x", "KeyRotateCW"
        '.OnKey "c", "KeyRotateCCW"
        
    End With
    On Error GoTo 0
End Sub

Sub RemoveKeyAssignations_2p()
    On Error Resume Next
    With Application
        'first player
        .OnKey "{DOWN}", ""
        .OnKey "{LEFT}", ""
        .OnKey "{RIGHT}", ""
        .OnKey "{UP}", ""
        .OnKey "m", ""
        
        'second player
        .OnKey "s", ""
        .OnKey "a", ""
        .OnKey "d", ""
        .OnKey "w", ""
        .OnKey "{TAB}", ""
    End With
    On Error GoTo 0
End Sub

Sub KeyDrop_2p()
    If GamSta_2p = 1 Then
        Call MoveBlockDown_2p(1)
    End If
End Sub

Sub KeyLeft_2p()
    If GamSta_2p = 1 Then
        Call MoveBlockRightLeft_2p(-1)
    End If
End Sub

Sub KeyRight_2p()
    If GamSta_2p = 1 Then
        Call MoveBlockRightLeft_2p(1)
    End If
End Sub

Sub KeyRotateCW_2p()
    If GamSta_2p = 1 Then
        Call RotateBlock_2p(1)
    End If
End Sub

Sub KeyRotateCCW_2p()
    If GamSta_2p = 1 Then
        Call RotateBlock_2p(-1)
    End If
End Sub

