Attribute VB_Name = "ModGameLogic"
Sub StartNewGame()
    'Switch Keyboard input to EN(US) to avoid game crashs by using Chinese Bopomofo input
    Call SwitchToEnglish
    ' Call the NewGame subroutine to start a new game
    Call NewGame
End Sub

Sub UpdateGame()
    ' Initialize game by setting initial values, preparing the game, creating the sheet
    Call SetInitialValues
    Call InitializeGame
    Call CreateGameSheet
End Sub

Function AddBlock(X As Byte, Y As Byte, Tem As Byte)
    ' Initialize variables for scoring and gap detection
    Dim Gap As Byte
    Dim Sco As Integer
    
    ' Loop through the size of the current block
    For i = 1 To CurBlo.Siz
        For j = 1 To CurBlo.Siz
            ' Check if there's a part of the block to add
            If CurBlo.Arr(1 + i, 1 + j) = 1 Then
                ' If temporary (Tem=1), place block part without scoring
                If Tem = 1 Then
                    Mat(X + i - 1, Y + j - 1) = 255
                Else
                    ' Otherwise, place block part and set color index from current block
                    Mat(X + i - 1, Y + j - 1) = CurBlo.ColInd
                    
                    ' Calculate score based on adjacency of placed parts to existing blocks
                    If CurBlo.Arr(0 + i, 1 + j) = 0 And Mat(X + i - 2, Y + j - 1) > 0 Then Sco = Sco + 10
                    If CurBlo.Arr(1 + i, 0 + j) = 0 And Mat(X + i - 1, Y + j - 2) > 0 Then Sco = Sco + 10
                    If CurBlo.Arr(2 + i, 1 + j) = 0 And Mat(X + i - 0, Y + j - 1) > 0 Then Sco = Sco + 10
                    If CurBlo.Arr(1 + i, 2 + j) = 0 And Mat(X + i - 1, Y + j - 0) > 0 Then Sco = Sco + 10

                    ' Calculate gaps for potential penalties
                    If CurBlo.Arr(1 + i, 0 + j) = 0 And Mat(X + i - 1, Y + j - 2) = 0 And Mat(X + 1 - 2, Y + j - 2) > 0 Then Gap = 1
                    If CurBlo.Arr(1 + i, 2 + j) = 0 And Mat(X + i - 1, Y + j - 0) = 0 And Mat(X + 1 - 2, Y + j - 0) > 0 Then Gap = 1
                    If CurBlo.Arr(2 + i, 1 + j) = 0 And Mat(X + i - 0, Y + j - 1) = 0 Then Gap = 1
                End If
            End If
        Next j
    Next i
    ' Update the game state with new score and gap status
    Sta.Sco = Sta.Sco + (Sta.Lev * Sco)
    Sta.GapSum = Sta.GapSum + Gap
    Sta.Gap = 1 - (Sta.GapSum / (Sta.Blo + 1))
End Function

Function GenerateBlocks(Blo As Byte)
    ' Initialize next block variable and randomization factor
    Dim Nex As Byte
    Dim Ran As Byte

    ' Debug: Print current block set and bounds
    Debug.Print "CurBloSet: "; CurBloSet
    Debug.Print "UBound(BloSet): "; UBound(BloSet)
    If CurBloSet > UBound(BloSet) Then
        MsgBox "Error: CurBloSet out of bounds", vbCritical
        Exit Function
    End If

    ' Determine the range of possible blocks
    If IsArray(BloSet(CurBloSet).Blo) Then
        Ran = UBound(BloSet(CurBloSet).Blo)
    Else
        MsgBox "Error: Blo array not initialized", vbCritical
        Exit Function
    End If

    ' Generate the next sequence of blocks
    For i = 1 To Blo
        Nex = NexBlo(1)
        NexBlo(1) = NexBlo(2)
        NexBlo(2) = NexBlo(3)
        NexBlo(3) = Int((Rnd * Ran) + 1)
    Next i

    ' Set up the next block with predefined settings
    With CurBlo
        .ColInd = Nex
        .NorCol = ChangeBrightness(ColLib(ColSet(CurColSet, .ColInd)).Nor, 20)
        .BriCol = ChangeBrightness(.NorCol, 240)
        .DarCol = ChangeBrightness(.NorCol, -40)
        .Siz = BloLib(BloSet(CurBloSet).Blo(Nex)).Siz
        .X = 4
        .Y = Int(PlaFie.W / 2 - .Siz / 2) + 4
    End With
    Call CopyBloLibArrToCurBloArr(Nex)
    If CurBlo.Arr(2, 2) + CurBlo.Arr(2, 3) + CurBlo.Arr(2, 4) + CurBlo.Arr(2, 5) = 0 Then CurBlo.X = 3
    If IsBlock(CurBlo.X, CurBlo.Y) = 0 Then
        Call Gameover
    Else
        Call AddBlock(CurBlo.X, CurBlo.Y, 1)
        If Blo = 1 Then Sta.Blo = Sta.Blo + 1
        Call DrawPlayingField(1)
        Call DisplayNextBlocks
        Call DisplayStatistics
    End If

    ' Adjust timers based on level
    If Tim.LevTim <> 17 - Sta.Lev Then
        Tim.LevTim = 17 - Sta.Lev
        Tim.ExeThr = Tim.LevTim
    End If
    Tim.CurPas = 0
End Function

Function IsBlock(X As Byte, Y As Byte) As Byte
    ' Check if the specified position on the matrix is occupied
    For i = 1 To CurBlo.Siz
        For j = 1 To CurBlo.Siz
            If CurBlo.Arr(1 + i, 1 + j) = 1 Then
                If Mat(X + i - 1, Y + j - 1) > 0 Then
                    IsBlock = 0
                    Exit Function
                End If
            End If
        Next j
    Next i
    IsBlock = 1
End Function

Function RemoveBlock(X As Byte, Y As Byte)
    ' Iterate through the size of the current block
    For i = 1 To CurBlo.Siz
        For j = 1 To CurBlo.Siz
            ' Check if the current cell in the block array is occupied (value = 1)
            If CurBlo.Arr(1 + i, 1 + j) = 1 Then
                ' Set the corresponding cell in the game matrix to 0 (empty)
                Mat(X + i - 1, Y + j - 1) = 0
            End If
        Next j
    Next i
End Function

Function MoveBlockDown(Dro As Byte)
    ' Attempt to move the block down, quickly if Dro = 1
    Call RemoveBlock(CurBlo.X, CurBlo.Y)
    If Dro = 1 Then
        ' Drop the block as far as possible
        While IsBlock(CurBlo.X + 1, CurBlo.Y) = 1
            CurBlo.X = CurBlo.X + 1
        Wend
        ' Drop the block as far as possible
        If IsBlock(CurBlo.X, CurBlo.Y - 1) = 1 And IsBlock(CurBlo.X - 1, CurBlo.Y - 1) = 0 Or _
           IsBlock(CurBlo.X, CurBlo.Y + 1) = 1 And IsBlock(CurBlo.X - 1, CurBlo.Y + 1) = 0 Then
            Call AddBlock(CurBlo.X, CurBlo.Y, 1)
            Call DrawPlayingField(0)
            Tim.CurPas = 0
        Else
            Call AddBlock(CurBlo.X, CurBlo.Y, 0)
            Call DrawPlayingField(0)
            If DeleteRows() = 0 Then Call GenerateBlocks(1)
        End If
    Else
        ' Move the block down by one row
        If IsBlock(CurBlo.X + 1, CurBlo.Y) = 1 Then
            CurBlo.X = CurBlo.X + 1
            Call AddBlock(CurBlo.X, CurBlo.Y, 1)
            Call DrawPlayingField(0)
        Else
            ' Handle block placement and row deletion if it can't move down further
            Call AddBlock(CurBlo.X, CurBlo.Y, 0)
            Call DrawPlayingField(0)
            If DeleteRows() = 0 Then Call GenerateBlocks(1)
        End If
    End If
End Function

Function MoveBlockRightLeft(Dir As Integer)
    ' Move the block to the right or left based on the direction
    Call RemoveBlock(CurBlo.X, CurBlo.Y)
    If IsBlock(CurBlo.X, CurBlo.Y + Dir) = 1 Then
        CurBlo.Y = CurBlo.Y + Dir
    End If
    Call AddBlock(CurBlo.X, CurBlo.Y, 1)
    Call DrawPlayingField(0)
End Function


Function RotateBlock(Rot As Integer)
    ' Rotate the block either clockwise or counterclockwise
    ReDim ArrCop(CurBlo.Siz, CurBlo.Siz) As Byte
    ReDim TemArr(CurBlo.Siz, CurBlo.Siz) As Byte
    Dim Siz As Byte
    
    Siz = CurBlo.Siz
    Call RemoveBlock(CurBlo.X, CurBlo.Y)
    ' Copy the block array for manipulation
    For i = 1 To Siz
        For j = 1 To Siz
            ArrCop(i, j) = CurBlo.Arr(1 + i, 1 + j)
        Next j
    Next i
    ' Rotate the block based on the direction
    For i = 1 To Siz
        For j = 1 To Siz
            If Rot = 1 Then
                CurBlo.Arr(1 + i, 1 + j) = ArrCop(Siz + 1 - j, i)
            Else
                CurBlo.Arr(1 + i, 1 + j) = ArrCop(j, Siz + 1 - i)
            End If
        Next j
    Next i
    ' Check block positioning and adjust if necessary
    If IsBlock(CurBlo.X, CurBlo.Y) = 1 Then
        Call AddBlock(CurBlo.X, CurBlo.Y, 1)
    Else
        If IsBlock(CurBlo.X, CurBlo.Y + 1) = 1 Then
            Call AddBlock(CurBlo.X, CurBlo.Y + 1, 1)
            CurBlo.Y = CurBlo.Y + 1
        Else
            If IsBlock(CurBlo.X, CurBlo.Y - 1) = 1 Then
                Call AddBlock(CurBlo.X, CurBlo.Y - 1, 1)
                CurBlo.Y = CurBlo.Y - 1
            Else
                If IsBlock(CurBlo.X + 1, CurBlo.Y) = 1 Then
                    Call AddBlock(CurBlo.X + 1, CurBlo.Y, 1)
                    CurBlo.X = CurBlo.X + 1
                Else
                    For i = 1 To Siz
                        For j = 1 To Siz
                             CurBlo.Arr(1 + i, 1 + j) = ArrCop(i, j)
                        Next j
                    Next i
                End If
                Call AddBlock(CurBlo.X, CurBlo.Y, 1)
            End If
        End If
    End If
    Call DrawPlayingField(0)
    
End Function

Function DeleteRows() As Byte
    ' Initialize counters for cells and completed rows
    Dim CelCou As Byte
    Dim RowCou As Byte
    
    ' Loop through each row within the game field boundaries
    For i = 4 To PlaFie.H + 3
        CelCou = 0  ' Reset cell counter for the new row
        ' Count filled cells in the current row
        For j = 4 To PlaFie.W + 3
            If Mat(i, j) > 0 Then CelCou = CelCou + 1
        Next j
        ' If all cells in the row are filled, clear the row
        If CelCou = PlaFie.W Then
            For j = 4 To PlaFie.W + 3
                Mat(i, j) = 0  ' Clear each cell in the row
            Next j
            RowCou = RowCou + 1  ' Increment the row counter
        End If
    Next i
    ' If any rows were cleared, update game state and timer settings
    If RowCou > 0 Then
        GamSta = 2  ' Change game state to indicate row clearing
        Tim.ExeThr = 5  ' Set execution threshold for timing events
        Tim.CurPas = 0  ' Reset the current pass counter
    End If
    ' Update statistics based on rows cleared
    Sta.Row = Sta.Row + RowCou
    If RowCou = 4 Then  ' Special case for clearing four rows (a Tetris)
        Sta.Qua = Sta.Qua + 1  ' Increment the Tetris count
        Sta.Sco = Sta.Sco + (RowCou * (Sta.Lev * 1000))  ' Score multiplier for a Tetris
        Sta.LevPro = Sta.LevPro + 12  ' Increment level progress significantly
    Else
        Sta.Sco = Sta.Sco + (RowCou * (Sta.Lev * 100))  ' Standard score increment
        Sta.LevPro = Sta.LevPro + (RowCou * 2)  ' Standard level progress increment
    End If
    Call CheckLevelProgress  ' Check and update level progress based on score and cleared rows
    DeleteRows = RowCou  ' Return the number of rows deleted
End Function

Sub DropRows()
    ' Loop through rows in the playing field
    For i = 4 To PlaFie.H + 3
        ' Check if the current row needs to be dropped (if it differs from the matrix copy)
        If Mat(i, 4) <> MatCop(i, 4) Then
            ' Shift all rows above down by one
            For j = i To 5 Step -1
                For k = 4 To PlaFie.W + 3
                    Mat(j, k) = Mat(j - 1, k)
                Next k
            Next j
            ' Clear the top row and update the matrix copy
            For j = 4 To PlaFie.W + 3
                Mat(4, j) = 0
                MatCop(i, j) = 0
            Next j
        End If
    Next i
    Call DrawPlayingField(1)  ' Redraw the playing field with updated rows
End Sub

Sub CheckLevelProgress()
    ' Check if the level progress has reached or exceeded 100 points
    If Sta.LevPro >= 100 Then
        Sta.LevPro = Sta.LevPro - 100  ' Reset level progress after incrementing level
        ' Check if the level is below the maximum level cap
        If Sta.Lev < 15 Then
            Sta.Lev = Sta.Lev + 1  ' Increment the level
            CurColSet = CurColSet + 1  ' Move to the next color set for blocks
        End If
    End If
End Sub

Sub DisplayGameoverInfo()
    
    Call SwitchToChineseBopomofo
    
    GGForm.ScoreLabel.Caption = " Your Score :    " & CStr(Sta.Sco)
    GGForm.MaxLabel.Caption = " Highest Score : " & CStr(Sta.ScoMax)
    GGForm.Show
   
End Sub
Sub Gameover()

    Call DrawPlayingField(0)
    Call EndTimer
    Call RemoveKeyAssignations
    GamSta = 5
    Call DisplayGameoverInfo
    
End Sub

Sub QuitTheGame()
    'pause game here
    response = MsgBox("Are you sure to quit this game?", vbOKCancel, "GameOver")
    
    If response = VbMsgBoxResult.vbOK Then
        Call Gameover
    'else
        'resume game
    End If
    
End Sub
