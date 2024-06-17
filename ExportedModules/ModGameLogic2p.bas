Attribute VB_Name = "ModGameLogic2p"
Sub StartNewGame_2p()
    'Switch Keyboard input to EN(US) to avoid game crashs by using Chinese Bopomofo input
    Call SwitchToEnglish
    ' Call the NewGame subroutine to start a new game
    Call Init_GameSheetSize_2p
    Call SetInitialValues
    Call InitializeGame_2p
    Call CreateGameSheet_2p
    Call NewGame_2p
End Sub

Sub UpdateGame_2p()
    ' Initialize game by setting initial values, preparing the game, creating the sheet
    Call SetInitialValues
    Call InitializeGame_2p
    Call CreateGameSheet_2p
End Sub

Function AddBlock_2p(X As Byte, Y As Byte, Tem As Byte)
    ' Initialize variables for scoring and gap detection
    Dim Gap_2p As Byte
    Dim Sco_2p As Integer
    
    ' Loop through the size of the current block
    For i = 1 To CurBlo_2p.Siz
        For j = 1 To CurBlo_2p.Siz
            ' Check if there's a part of the block to add
            If CurBlo_2p.Arr(1 + i, 1 + j) = 1 Then
                ' If temporary (Tem=1), place block part without scoring
                If Tem = 1 Then
                    Mat_2p(X + i - 1, Y + j - 1) = 255
                Else
                    ' Otherwise, place block part and set color index from current block
                    Mat_2p(X + i - 1, Y + j - 1) = CurBlo_2p.ColInd
                    
                    ' Calculate score based on adjacency of placed parts to existing blocks
                    'If CurBlo.Arr(0 + i, 1 + j) = 0 And Mat(X + i - 2, Y + j - 1) > 0 Then Sco = Sco + 10
                    'If CurBlo.Arr(1 + i, 0 + j) = 0 And Mat(X + i - 1, Y + j - 2) > 0 Then Sco = Sco + 10
                    'If CurBlo.Arr(2 + i, 1 + j) = 0 And Mat(X + i - 0, Y + j - 1) > 0 Then Sco = Sco + 10
                    'If CurBlo.Arr(1 + i, 2 + j) = 0 And Mat(X + i - 1, Y + j - 0) > 0 Then Sco = Sco + 10

                    ' Calculate gaps for potential penalties
                    If CurBlo_2p.Arr(1 + i, 0 + j) = 0 And Mat_2p(X + i - 1, Y + j - 2) = 0 And Mat_2p(X + 1 - 2, Y + j - 2) > 0 Then Gap_2p = 1
                    If CurBlo_2p.Arr(1 + i, 2 + j) = 0 And Mat_2p(X + i - 1, Y + j - 0) = 0 And Mat_2p(X + 1 - 2, Y + j - 0) > 0 Then Gap_2p = 1
                    If CurBlo_2p.Arr(2 + i, 1 + j) = 0 And Mat_2p(X + i - 0, Y + j - 1) = 0 Then Gap_2p = 1
                End If
            End If
        Next j
    Next i
    ' Update the game state with new score and gap status
    Sta_2p.Sco = Sta_2p.Sco + (Sta_2p.Lev * Sco_2p)
    Sta_2p.GapSum = Sta_2p.GapSum + Gap_2p
    Sta_2p.Gap = 1 - (Sta_2p.GapSum / (Sta_2p.Blo + 1))
End Function

Function GenerateBlocks_2p(Blo_2p As Byte)
    ' Initialize next block variable and randomization factor
    Dim Nex_2p As Byte
    Dim Ran_2p As Byte

    ' Debug: Print current block set and bounds
    Debug.Print "CurBloSet: "; CurBloSet
    Debug.Print "UBound(BloSet): "; UBound(BloSet)
    If CurBloSet > UBound(BloSet) Then
        MsgBox "Error: CurBloSet out of bounds", vbCritical
        Exit Function
    End If

    ' Determine the range of possible blocks
    If IsArray(BloSet(CurBloSet).Blo) Then
        Ran_2p = UBound(BloSet(CurBloSet).Blo)
    Else
        MsgBox "Error: Blo array not initialized", vbCritical
        Exit Function
    End If

    ' Generate the next sequence of blocks
    For i = 1 To Blo_2p
        Nex_2p = NexBlo_2p(1)
        NexBlo_2p(1) = NexBlo_2p(2)
        NexBlo_2p(2) = NexBlo_2p(3)
        NexBlo_2p(3) = Int((Rnd * Ran_2p) + 1)
    Next i

    ' Set up the next block with predefined settings
    With CurBlo_2p
        .ColInd = Nex_2p
        .NorCol = ChangeBrightness(ColLib(ColSet(CurColSet, .ColInd)).Nor, 20)
        .BriCol = ChangeBrightness(.NorCol, 240)
        .DarCol = ChangeBrightness(.NorCol, -40)
        .Siz = BloLib(BloSet(CurBloSet).Blo(Nex_2p)).Siz
        .X = 4
        .Y = Int(PlaFie_2p.W / 2 - .Siz / 2) + PFY_2p + 1
    End With
    Call CopyBloLibArrToCurBloArr_2p(Nex_2p)
    If CurBlo_2p.Arr(2, 2) + CurBlo_2p.Arr(2, 3) + CurBlo_2p.Arr(2, 4) + CurBlo_2p.Arr(2, 5) = 0 Then CurBlo_2p.X = 3
    If IsBlock_2p(CurBlo_2p.X, CurBlo_2p.Y) = 0 Then
        Call Gameover_2p
        ' Update history list
        'Call UpdateGameRecord(Sta.Sco, Sta.Lev, Sta.Row, Sta.Qua)
    Else
        'If IsGamePaused = False Then
            Call AddBlock_2p(CurBlo_2p.X, CurBlo_2p.Y, 1)
        'End If
        If Blo_2p = 1 Then Sta_2p.Blo = Sta_2p.Blo + 1
        Call DrawPlayingField_2p(1)
        Call DisplayStatistics_2p
    End If

    ' Adjust timers based on level
    If Tim_2p.LevTim <> 17 - Sta_2p.Lev Then
        Tim_2p.LevTim = 17 - Sta_2p.Lev
        Tim_2p.ExeThr = Tim_2p.LevTim
    End If
    Tim_2p.CurPas = 0
End Function

Function IsBlock_2p(X As Byte, Y As Byte) As Byte
    ' Check if the specified position on the matrix is occupied
    For i = 1 To CurBlo_2p.Siz
        For j = 1 To CurBlo_2p.Siz
            If CurBlo_2p.Arr(1 + i, 1 + j) = 1 Then
                If Mat_2p(X + i - 1, Y + j - 1) > 0 Then
                    IsBlock_2p = 0
                    Exit Function
                End If
            End If
        Next j
    Next i
    IsBlock_2p = 1
End Function

Function RemoveBlock_2p(X As Byte, Y As Byte)
    ' Iterate through the size of the current block
    For i = 1 To CurBlo_2p.Siz
        For j = 1 To CurBlo_2p.Siz
            ' Check if the current cell in the block array is occupied (value = 1)
            If CurBlo_2p.Arr(1 + i, 1 + j) = 1 Then
                ' Set the corresponding cell in the game matrix to 0 (empty)
                Mat_2p(X + i - 1, Y + j - 1) = 0
            End If
        Next j
    Next i
End Function

Function MoveBlockDown_2p(Dro_2p As Byte)
    ' Attempt to move the block down, quickly if Dro = 1
    Call RemoveBlock_2p(CurBlo_2p.X, CurBlo_2p.Y)
    If Dro_2p = 1 Then
        ' Drop the block as far as possible
        While IsBlock_2p(CurBlo_2p.X + 1, CurBlo_2p.Y) = 1
            CurBlo_2p.X = CurBlo_2p.X + 1
        Wend
        ' Drop the block as far as possible
        If IsBlock_2p(CurBlo_2p.X, CurBlo_2p.Y - 1) = 1 And IsBlock_2p(CurBlo_2p.X - 1, CurBlo_2p.Y - 1) = 0 Or _
           IsBlock_2p(CurBlo_2p.X, CurBlo_2p.Y + 1) = 1 And IsBlock_2p(CurBlo_2p.X - 1, CurBlo_2p.Y + 1) = 0 Then
            Call AddBlock_2p(CurBlo_2p.X, CurBlo_2p.Y, 1)
            Call DrawPlayingField_2p(0)
            Tim.CurPas = 0
        Else
            Call AddBlock_2p(CurBlo_2p.X, CurBlo_2p.Y, 0)
            Call DrawPlayingField_2p(0)
            If DeleteRows_2p() = 0 Then Call GenerateBlocks_2p(1)
        End If
    Else
        ' Move the block down by one row
        If IsBlock_2p(CurBlo_2p.X + 1, CurBlo_2p.Y) = 1 Then
            CurBlo_2p.X = CurBlo_2p.X + 1
            Call AddBlock_2p(CurBlo_2p.X, CurBlo_2p.Y, 1)
            Call DrawPlayingField_2p(0)
        Else
            ' Handle block placement and row deletion if it can't move down further
            Call AddBlock_2p(CurBlo_2p.X, CurBlo_2p.Y, 0)
            Call DrawPlayingField_2p(0)
            If DeleteRows_2p() = 0 Then Call GenerateBlocks_2p(1)
        End If
    End If
End Function

Function MoveBlockRightLeft_2p(Dir_2p As Integer)
    ' Move the block to the right or left based on the direction
    Call RemoveBlock_2p(CurBlo_2p.X, CurBlo_2p.Y)
    If IsBlock_2p(CurBlo_2p.X, CurBlo_2p.Y + Dir_2p) = 1 Then
        CurBlo.Y = CurBlo.Y + Dir
    End If
    Call AddBlock_2p(CurBlo_2p.X, CurBlo_2p.Y, 1)
    Call DrawPlayingField_2p(0)
End Function


Function RotateBlock_2p(Rot_2p As Integer)
    ' Rotate the block either clockwise or counterclockwise
    ReDim ArrCop_2p(CurBlo_2p.Siz, CurBlo_2p.Siz) As Byte
    ReDim TemArr_2p(CurBlo_2p.Siz, CurBlo_2p.Siz) As Byte
    Dim Siz_2p As Byte
    
    Siz_2p = CurBlo_2p.Siz
    Call RemoveBlock_2p(CurBlo_2p.X, CurBlo_2p.Y)
    ' Copy the block array for manipulation
    For i = 1 To Siz_2p
        For j = 1 To Siz_2p
            ArrCop_2p(i, j) = CurBlo_2p.Arr(1 + i, 1 + j)
        Next j
    Next i
    ' Rotate the block based on the direction
    For i = 1 To Siz_2p
        For j = 1 To Siz_2p
            If Rot_2p = 1 Then
                CurBlo_2p.Arr(1 + i, 1 + j) = ArrCop_2p(Siz_2p + 1 - j, i)
            Else
                CurBlo_2p.Arr(1 + i, 1 + j) = ArrCop_2p(j, Siz_2p + 1 - i)
            End If
        Next j
    Next i
    ' Check block positioning and adjust if necessary
    If IsBlock_2p(CurBlo_2p.X, CurBlo_2p.Y) = 1 Then
        Call AddBlock_2p(CurBlo_2p.X, CurBlo_2p.Y, 1)
    Else
        If IsBlock_2p(CurBlo_2p.X, CurBlo_2p.Y + 1) = 1 Then
            Call AddBlock_2p(CurBlo_2p.X, CurBlo_2p.Y + 1, 1)
            CurBlo_2p.Y = CurBlo_2p.Y + 1
        Else
            If IsBlock_2p(CurBlo_2p.X, CurBlo_2p.Y - 1) = 1 Then
                Call AddBlock_2p(CurBlo_2p.X, CurBlo_2p.Y - 1, 1)
                CurBlo_2p.Y = CurBlo_2p.Y - 1
            Else
                If IsBlock_2p(CurBlo_2p.X + 1, CurBlo_2p.Y) = 1 Then
                    Call AddBlock_2p(CurBlo_2p.X + 1, CurBlo_2p.Y, 1)
                    CurBlo_2p.X = CurBlo_2p.X + 1
                Else
                    For i = 1 To Siz_2p
                        For j = 1 To Siz_2p
                             CurBlo_2p.Arr(1 + i, 1 + j) = ArrCop_2p(i, j)
                        Next j
                    Next i
                End If
                Call AddBlock_2p(CurBlo_2p.X, CurBlo_2p.Y, 1)
            End If
        End If
    End If
    Call DrawPlayingField_2p(0)
    
End Function

Function DeleteRows_2p() As Byte
    ' Initialize counters for cells and completed rows
    Dim CelCou_2p As Byte
    Dim RowCou_2p As Byte
    
    ' Loop through each row within the game field boundaries
    For i = 4 To PlaFie_2p.H + 3
        CelCou = 0  ' Reset cell counter for the new row
        ' Count filled cells in the current row
        For j = 4 To PlaFie_2p.W + 3
            If Mat_2p(i, j) > 0 Then CelCou_2p = CelCou_2p + 1
        Next j
        ' If all cells in the row are filled, clear the row
        If CelCou_2p = PlaFie_2p.W Then
            For j = 4 To PlaFie_2p.W + 3
                Mat_2p(i, j) = 0  ' Clear each cell in the row
            Next j
            RowCou_2p = RowCou_2p + 1  ' Increment the row counter
        End If
    Next i
    ' If any rows were cleared, update game state and timer settings
    If RowCou_2p > 0 Then
        GamSta_2p = 2  ' Change game state to indicate row clearing
        Tim_2p.ExeThr = 5  ' Set execution threshold for timing events
        Tim_2p.CurPas = 0  ' Reset the current pass counter
    End If
    ' Update statistics based on rows cleared
    Sta_2p.Row = Sta_2p.Row + RowCou_2p
    If RowCou_2p = 4 Then  ' Special case for clearing four rows (a Tetris)
        Sta_2p.Qua = Sta_2p.Qua + 1  ' Increment the Tetris count
        Sta_2p.Sco = Sta_2p.Sco + (RowCou_2p * (Sta_2p.Lev * 1000))  ' Score multiplier for a Tetris
        Sta_2p.LevPro = Sta_2p.LevPro + 12  ' Increment level progress significantly
    Else
        Sta_2p.Sco = Sta_2p.Sco + (RowCou_2p * (Sta_2p.Lev * 100))  ' Standard score increment
        Sta_2p.LevPro = Sta_2p.LevPro + (RowCou_2p * 2)  ' Standard level progress increment
    End If
    Call CheckLevelProgress_2p  ' Check and update level progress based on score and cleared rows
    DeleteRows_2p = RowCou_2p  ' Return the number of rows deleted
End Function

Sub DropRows_2p()
    ' Loop through rows in the playing field
    For i = 4 To PlaFie_2p.H + 3
        ' Check if the current row needs to be dropped (if it differs from the matrix copy)
        If Mat_2p(i, 4) <> MatCop_2p(i, 4) Then
            ' Shift all rows above down by one
            For j = i To 5 Step -1
                For k = 4 To PlaFie_2p.W + 3
                    Mat_2p(j, k) = Mat_2p(j - 1, k)
                Next k
            Next j
            ' Clear the top row and update the matrix copy
            For j = 4 To PlaFie_2p.W + 3
                Mat_2p(4, j) = 0
                MatCop_2p(i, j) = 0
            Next j
        End If
    Next i
    Call DrawPlayingField_2p(1)  ' Redraw the playing field with updated rows
End Sub

Sub CheckLevelProgress_2p()
    ' Check if the level progress has reached or exceeded 100 points
    If Sta_2p.LevPro >= 100 Then
        Sta_2p.LevPro = Sta_2p.LevPro - 100  ' Reset level progress after incrementing level
        ' Check if the level is below the maximum level cap
        If Sta_2p.Lev < 15 Then
            Sta_2p.Lev = Sta_2p.Lev + 1  ' Increment the level
            CurColSet = CurColSet + 1  ' Move to the next color set for blocks
        End If
    End If
End Sub

Sub Gameover_2p()

    Call DrawPlayingField(0)
    Call DrawPlayingField_2p(0)
    Call EndTimer
    Call EndTimer_2p
    Call RemoveKeyAssignations
    Call RemoveKeyAssignations_2p
    GamSta = 5
    GamSta_2p = 5
    'Call DisplayGameoverInfo

End Sub
