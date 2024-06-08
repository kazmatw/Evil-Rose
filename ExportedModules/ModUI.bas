Attribute VB_Name = "ModUI"
Sub CreateGameSheet()
    ' Declare variables for colors, dimensions, and positions of the game and statistics fields
    Dim PFBBC, PFBNC, PFBDC As Long  ' Game field border colors
    Dim PFBC1, PFBC2 As Long  ' Game field background colors
    Dim PFH, PFW As Byte  ' Game field height and width
    Dim PFX, PFY As Byte  ' Game field position
    Dim SFBBC, SFBNC, SFBDC As Long  ' Statistics field border colors
    Dim SFBC1, SFBC2 As Long  ' Statistics field background colors
    Dim SFH, SFW As Byte  ' Statistics field height and width
    Dim SFX, SFY As Byte  ' Statistics field position
    
    ' Load game field properties from PlaFie
    With PlaFie
        PFBC1 = .BacCol1
        PFBC2 = .BacCol2
        PFBBC = .BorBCol
        PFBDC = .BorDCol
        PFBNC = .BorNCol
        PFH = .H
        PFW = .W
        PFX = .X
        PFY = .Y
    End With
    
    ' Load statistics field properties from StaFie
    With StaFie
        SFBC1 = .BacCol1
        SFBC2 = .BacCol2
        SFBBC = .BorBCol
        SFBDC = .BorDCol
        SFBNC = .BorNCol
        .H = PFH
        .Y = PFY + PFW + 3
        SFW = .W
        SFH = .H
        SFX = .X
        SFY = .Y
    End With
        
    ' Reset cells in the range A1:AF26
    With Range("A1:AF26")
        .Value = ""  ' Clear cell values
        .Borders.LineStyle = -4142  ' Remove cell borders
        .UnMerge  ' Unmerge any merged cells
    End With
    
    ' Draw game sheet background
    Range(Cells(PFX - 2, PFY - 2), Cells(PFX - 2, PFY + PFW + SFW + 4)).Interior.Color = GamSheBC
    Range(Cells(PFX + PFH + 1, PFY - 2), Cells(PFX + PFH + 1, PFY + PFW + SFW + 4)).Interior.Color = GamSheBC
    Range(Cells(PFX + PFH + 2, PFY - 2), Cells(PFX + PFH + 12, PFY + PFW + SFW + 10)).Interior.Color = xlNone
    Range(Cells(PFX - 2, PFY - 2), Cells(PFX + PFH + 1, PFY - 2)).Interior.Color = GamSheBC
    Range(Cells(PFX - 2, PFY + PFW + 1), Cells(PFX + PFH + 1, PFY + PFW + 1)).Interior.Color = GamSheBC
    Range(Cells(PFX - 2, PFY + PFW + SFW + 4), Cells(PFX + PFH + 1, PFY + PFW + SFW + 4)).Interior.Color = GamSheBC
    Range(Cells(PFX - 2, PFY + PFW + SFW + 5), Cells(PFX + PFH + 1, PFY + PFW + SFW + 10)).Interior.Color = xlNone
   
    ' Draw game field with borders and background color
    With Range(Cells(PFX - 1, PFY - 1), Cells(PFX + PFH, PFY + PFW))
        .Borders(8).Color = PFBBC  ' Top border
        .Borders(8).Weight = 4
        .Borders(7).Color = PFBBC  ' Bottom border
        .Borders(7).Weight = 4
        .Borders(9).Color = PFBDC  ' Left border
        .Borders(9).Weight = 4
        .Borders(10).Color = PFBDC  ' Right border
        .Borders(10).Weight = 4
    End With
    ' Set interior color for the borders
    Range(Cells(PFX - 1, PFY - 1), Cells(PFX + PFH, PFY + PFW)).Interior.Color = PFBNC
    ' Set interior properties for the playing field cells
    With Range(Cells(PFX, PFY), Cells(PFX + PFH - 1, PFY + PFW - 1))
        .HorizontalAlignment = 3  ' Center horizontally
        .VerticalAlignment = 2  ' Center vertically
        .Font.Color = PFBC2
        .Font.Name = "Arial"
        .Font.Bold = 1
        .Font.Size = 24
        .Value = "X"  ' Placeholder value for testing
        .Interior.Color = PFBC1
        .Borders(8).Color = PFBDC
        .Borders(8).Weight = 4
        .Borders(7).Color = PFBDC
        .Borders(7).Weight = 4
        .Borders(9).Color = PFBBC
        .Borders(9).Weight = 4
        .Borders(10).Color = PFBBC
        .Borders(10).Weight = 4
    End With
    
    ' Draw statistics field with borders and background color
    Range(Cells(SFX, SFY), Cells(SFX, SFY + 5)).Merge  ' Merge cells for "NEXT" label
    For i = SFX To SFX + SFH - 1
        Range(Cells(i, SFY + 7), Cells(i, SFY + 12)).Merge  ' Merge cells for score display
    Next i
    With Range(Cells(SFX - 1, SFY - 1), Cells(SFX + SFH, SFY + SFW))
        .Borders(8).Color = SFBBC  ' Top border
        .Borders(8).Weight = 4
        .Borders(7).Color = SFBBC  ' Bottom border
        .Borders(7).Weight = 4
        .Borders(9).Color = SFBDC  ' Left border
        .Borders(9).Weight = 4
        .Borders(10).Color = SFBDC  ' Right border
        .Borders(10).Weight = 4
    End With
    ' Set interior color for the borders
    Range(Cells(SFX - 1, SFY - 1), Cells(SFX - 1, SFY + SFW)).Interior.Color = SFBNC
    Range(Cells(SFX + SFH, SFY - 1), Cells(SFX + SFH, SFY + SFW)).Interior.Color = SFBNC
    Range(Cells(SFX - 1, SFY - 1), Cells(SFX + SFH, SFY - 1)).Interior.Color = SFBNC
    Range(Cells(SFX - 1, SFY + 6), Cells(SFX + SFH, SFY + 6)).Interior.Color = SFBNC
    Range(Cells(SFX - 1, SFY + SFW), Cells(SFX + SFH, SFY + SFW)).Interior.Color = SFBNC
    ' Set interior properties for the statistics field cells
    With Range(Cells(SFX, SFY), Cells(SFX + SFH - 1, SFY + 5))
        .Borders(8).Color = SFBDC
        .Borders(8).Weight = 4
        .Borders(7).Color = SFBDC
        .Borders(7).Weight = 4
        .Borders(9).Color = SFBBC
        .Borders(9).Weight = 4
        .Borders(10).Color = SFBBC
        .Borders(10).Weight = 4
        .Interior.Color = SFBC1
    End With
    ' Set interior properties for the score display cells
    With Range(Cells(SFX, SFY + 7), Cells(SFX + SFH - 1, SFY + 12))
        .Borders(8).Color = SFBDC
        .Borders(8).Weight = 4
        .Borders(7).Color = SFBDC
        .Borders(7).Weight = 4
        .Borders(9).Color = SFBBC
        .Borders(9).Weight = 4
        .Borders(10).Color = SFBBC
        .Borders(10).Weight = 4
        .Interior.Color = SFBC1
    End With
    ' Configure the "NEXT" label cell
    With Cells(SFX, SFY)
        .Font.Color = &H884444
        .Font.Bold = True
        .Font.Italic = True
        .Font.Name = "Arial"
        .Font.Size = 18
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .Value = "NEXT"
    End With
    ' Configure score and max score cells
    For i = SFX To SFX + 2 Step 2
        With Cells(i, SFY + 7)
            .Font.Color = &H884444
            .Font.Bold = True
            .Font.Italic = True
            .Font.Name = "Arial"
            .Font.Size = 18
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            If i = SFX Then
                .Value = "SCORE"
            Else
                .Value = "MAX SCORE"
            End If
        End With
        With Cells(i + 1, SFY + 7)
            .Font.Color = &HFFDDDD
            .Font.Bold = True
            .Font.Italic = False
            .Font.Name = "Arial"
            .Font.Size = 20
            .HorizontalAlignment = xlRight
            .IndentLevel = 4
            .VerticalAlignment = xlBottom
            .Value = 0
        End With
    Next i
    ' Configure remaining statistics cells
    For i = SFX + 5 To SFX + 13 Step 2
        With Cells(i, SFY + 7)
            .Font.Color = &H884444
            .Font.Bold = True
            .Font.Italic = True
            .Font.Name = "Arial"
            .Font.Size = 18
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            Select Case i
                Case SFX + 5
                    .Value = "LEVEL"
                Case SFX + 7
                    .Value = "BLOCKS"
                Case SFX + 9
                    .Value = "ROWS"
                Case SFX + 11
                    .Value = "QUADS"
                Case Else
                    .Value = "GAPLESS"
            End Select
        End With
        With Cells(i + 1, SFY + 7)
            .Font.Color = &HFF8888
            .Font.Bold = True
            .Font.Italic = False
            .Font.Name = "Arial"
            .Font.Size = 20
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .Value = 0
            If i = SFX + 13 Then
                .NumberFormat = "0%"
            Else
                .NumberFormat = ""
            End If
        End With
    Next i
    
End Sub

Sub DrawPlayingField(Mode As Byte)
    ' Declare static variable to hold the previous level
    Static PreLev As Byte
    ' Declare variables for coordinates and dimensions
    Dim X, Y As Byte
    Dim H, W As Byte
    
    ' Load game field properties from PlaFie
    With PlaFie
        X = .X  ' X coordinate of the game field
        Y = .Y  ' Y coordinate of the game field
        H = .H  ' Height of the game field
        W = .W  ' Width of the game field
    End With

    ' If Mode is greater than 0, reset the inner borders of the range (commented out in provided code)
'    If Mode > 0 Then
'        With Range(Cells(X, Y), Cells(X + H - 1, Y + W - 1))
'            .Borders(11).LineStyle = -4142  ' Diagonal down border
'            .Borders(12).LineStyle = -4142  ' Diagonal up border
'        End With
'    End If

    ' Loop through the game matrix dimensions
    For i = 4 To H + 3
        For j = 4 To W + 3
            ' Check if the current cell needs updating
            If Mat(i, j) <> MatCop(i, j) Or Mode = 1 And Sta.Lev <> PreLev Or Mode = 2 Then
                ' Update the cell in the game field
                With Cells(X + i - 4, Y + j - 4)
                    Select Case Mat(i, j)
                        ' Handle the case for current block (value 255)
                        Case 255
                            .Value = ""
                            .Interior.Color = CurBlo.NorCol  ' Set the normal color
                            ' Set the top border color and weight
                            If Mat(i - 1, j) < 255 Then
                                .Borders(8).Color = CurBlo.BriCol
                                .Borders(8).Weight = 4
                            Else
                                .Borders(8).LineStyle = -4142
                            End If
                            ' Set the bottom border color and weight
                            If Mat(i, j - 1) < 255 Then
                                .Borders(7).Color = CurBlo.BriCol
                                .Borders(7).Weight = 4
                            Else
                                .Borders(7).LineStyle = -4142
                            End If
                            ' Set the left border color and weight
                            If Mat(i + 1, j) < 255 Then
                                .Borders(9).Color = CurBlo.DarCol
                                .Borders(9).Weight = 4
                            Else
                                .Borders(9).LineStyle = -4142
                            End If
                            ' Set the right border color and weight
                            If Mat(i, j + 1) < 255 Then
                                .Borders(10).Color = CurBlo.DarCol
                                .Borders(10).Weight = 4
                            Else
                                .Borders(10).LineStyle = -4142
                            End If
                        ' Handle the case for normal block (value > 0)
                        Case Is > 0
                            .Value = ""
                            .Interior.Color = ColLib(ColSet(CurColSet, Mat(i, j))).Nor  ' Set the normal color
                            ' Set the top border color and weight
                            If i > 4 Then
                                If Mat(i - 1, j) = 0 Then
                                    .Borders(8).Color = ColLib(ColSet(CurColSet, Mat(i, j))).Bri
                                    .Borders(8).Weight = 2
                                ElseIf Mat(i - 1, j) <> Mat(i, j) And Mat(i - 1, j) < 255 Then
                                    .Borders(8).Color = &H444444
                                    .Borders(8).Weight = 2
                                Else
                                    .Borders(8).LineStyle = -4142
                                End If
                            Else
                                .Borders(8).Color = PlaFie.BorDCol
                                .Borders(8).Weight = 4
                            End If
                            ' Set the bottom border color and weight
                            If j > 4 Then
                                If Mat(i, j - 1) = 0 Then
                                    .Borders(7).Color = ColLib(ColSet(CurColSet, Mat(i, j))).Bri
                                    .Borders(7).Weight = 2
                                ElseIf Mat(i, j - 1) <> Mat(i, j) And Mat(i, j - 1) < 255 Then
                                    .Borders(7).Color = &H444444
                                    .Borders(7).Weight = 2
                                Else
                                    .Borders(7).LineStyle = -4142
                                End If
                            Else
                                .Borders(7).Color = PlaFie.BorDCol
                                .Borders(7).Weight = 4
                            End If
                            ' Set the left border color and weight
                            If i < H + 3 Then
                                If Mat(i + 1, j) = 0 Then
                                    .Borders(9).Color = ColLib(ColSet(CurColSet, Mat(i, j))).Dar
                                    .Borders(9).Weight = 2
                                ElseIf Mat(i + 1, j) <> Mat(i, j) And Mat(i + 1, j) < 255 Then
                                    .Borders(9).Color = &H444444
                                    .Borders(9).Weight = 2
                                Else
                                    .Borders(9).LineStyle = -4142
                                End If
                            Else
                                .Borders(9).Color = PlaFie.BorBCol
                                .Borders(9).Weight = 4
                            End If
                            ' Set the right border color and weight
                            If j < W + 3 Then
                                If Mat(i, j + 1) = 0 Then
                                    .Borders(10).Color = ColLib(ColSet(CurColSet, Mat(i, j))).Dar
                                    .Borders(10).Weight = 2
                                ElseIf Mat(i, j + 1) <> Mat(i, j) And Mat(i, j + 1) < 255 Then
                                    .Borders(10).Color = &H444444
                                    .Borders(10).Weight = 2
                                Else
                                    .Borders(10).LineStyle = -4142
                                End If
                            Else
                                .Borders(10).Color = PlaFie.BorBCol
                                .Borders(10).Weight = 4
                            End If
                        ' Handle the case for empty cell (value = 0)
                        Case Else
                            .Interior.Color = PlaFie.BacCol1  ' Set the background color
                            .Value = "T"  ' Placeholder value for testing
                            ' Set the top border color and weight
                            If i = 4 Then
                                .Borders(8).Color = PlaFie.BorDCol
                                .Borders(8).Weight = 4
                            ElseIf Mat(i - 1, j) = 255 Then
                                .Borders(8).Color = CurBlo.DarCol
                                .Borders(8).Weight = 4
                            ElseIf Mat(i - 1, j) > 0 Then
                                .Borders(8).Color = ColLib(ColSet(CurColSet, Mat(i - 1, j))).Dar
                                .Borders(8).Weight = 2
                            Else
                                .Borders(8).LineStyle = -4142
                            End If
                            ' Set the bottom border color and weight
                            If j = 4 Then
                                .Borders(7).Color = PlaFie.BorDCol
                                .Borders(7).Weight = 4
                            ElseIf Mat(i, j - 1) = 255 Then
                                .Borders(7).Color = CurBlo.DarCol
                                .Borders(7).Weight = 4
                            ElseIf Mat(i, j - 1) > 0 Then
                                .Borders(7).Color = ColLib(ColSet(CurColSet, Mat(i, j - 1))).Dar
                                .Borders(7).Weight = 2
                            Else
                                .Borders(7).LineStyle = -4142
                            End If
                            ' Set the left border color and weight
                            If i = H + 3 Then
                                .Borders(9).Color = PlaFie.BorBCol
                                .Borders(9).Weight = 4
                            ElseIf Mat(i + 1, j) = 255 Then
                                .Borders(9).Color = CurBlo.BriCol
                                .Borders(9).Weight = 4
                            ElseIf Mat(i + 1, j) > 0 Then
                                .Borders(9).Color = ColLib(ColSet(CurColSet, Mat(i + 1, j))).Bri
                                .Borders(9).Weight = 2
                            Else
                                .Borders(9).LineStyle = -4142
                            End If
                            ' Set the right border color and weight
                            If j = W + 3 Then
                                .Borders(10).Color = PlaFie.BorBCol
                                .Borders(10).Weight = 4
                            ElseIf Mat(i, j + 1) = 255 Then
                                .Borders(10).Color = CurBlo.BriCol
                                .Borders(10).Weight = 4
                            ElseIf Mat(i, j + 1) > 0 Then
                                .Borders(10).Color = ColLib(ColSet(CurColSet, Mat(i, j + 1))).Bri
                                .Borders(10).Weight = 2
                            Else
                                .Borders(10).LineStyle = -4142
                            End If
                    End Select
                End With
                ' Update the matrix copy to reflect the changes
                MatCop(i, j) = Mat(i, j)
            End If
        Next j
    Next i
    
    ' Store the current level as the previous level for the next update
    PreLev = Sta.Lev
End Sub

Sub DrawDeletedRows()
    ' Declare variables for coordinates of the game field
    Dim X, Y As Byte
    
    ' Load game field position from PlaFie
    X = PlaFie.X  ' X coordinate of the game field
    Y = PlaFie.Y  ' Y coordinate of the game field

    ' Loop through each row in the game matrix
    For i = 4 To PlaFie.H + 3
        ' Check if the current row has changed (indicating it was deleted)
        If Mat(i, 4) <> MatCop(i, 4) Then
            ' Update the cells in the row to reflect the deletion
            With Range(Cells(X + i - 4, Y), Cells(X + i - 4, Y + PlaFie.W - 1))
                .Borders.LineStyle = -4142  ' Remove all borders
                .Borders(7).Color = PlaFie.BorDCol  ' Set bottom border color
                .Borders(7).Weight = 4  ' Set bottom border weight
                .Borders(10).Color = PlaFie.BorBCol  ' Set right border color
                .Borders(10).Weight = 4  ' Set right border weight
                ' Set top border for the first row
                If i = 4 Then
                    .Borders(8).Color = PlaFie.BorDCol  ' Set top border color
                    .Borders(8).Weight = 4  ' Set top border weight
                End If
                ' Set bottom border for the last row
                If i = PlaFie.H + 3 Then
                    .Borders(9).Color = PlaFie.BorBCol  ' Set bottom border color
                    .Borders(9).Weight = 4  ' Set bottom border weight
                End If
                .Interior.Color = PlaFie.BacCol1  ' Set interior color
                .Value = "T"  ' Set placeholder value for testing
            End With
        End If
    Next i
    
    ' Update the game state to indicate rows have been deleted
    GamSta = 3
End Sub

Sub DisplayNextBlocks()
    ' Declare variables for the block array, block height, colors, and positioning
    Dim Arr(6, 6) As Byte  ' Array to store the block structure
    Dim BloHei As Byte  ' Variable to hold the height of the block
    Dim Dar, Nor, Bri As Long  ' Variables for dark, normal, and bright colors
    Dim RowVal As Byte  ' Variable to hold the value of the row
    Dim X, Y As Byte  ' Variables for coordinates
    
    ' Set the starting position for the display of the next blocks
    X = StaFie.X + 2
    Y = StaFie.Y + 1
    
    ' Clear the area where the next blocks will be displayed
    With Range(Cells(X, Y), Cells(X + 10, Y + 3))
        .Borders.LineStyle = -4142  ' Remove borders
        .Interior.Color = StaFie.BacCol1  ' Set background color
    End With
    
    ' Loop through the number of next blocks to display
    For i = 1 To BloPre
        ' Get the colors for the next block
        Dar = ColLib(ColSet(CurColSet, NexBlo(i))).Dar
        Nor = ColLib(ColSet(CurColSet, NexBlo(i))).Nor
        Bri = ColLib(ColSet(CurColSet, NexBlo(i))).Bri
        
        ' Copy the block structure to the array
        For j = 1 To 4
            For k = 1 To 4
                Arr(1 + j, 1 + k) = BloLib(BloSet(CurBloSet).Blo(NexBlo(i))).Arr(1 + j, 1 + k)
            Next k
        Next j
        
        BloHei = 0  ' Reset the block height
        
        ' Adjust the X position and block height if the block is on the second row
        If Arr(2, 2) + Arr(2, 3) + Arr(2, 4) + Arr(2, 5) = 0 Then
            X = X - 1
            BloHei = 1
        End If
        
        ' Loop through the block structure to set cell properties
        For j = 1 To 4
            RowVal = 0  ' Reset the row value
            For k = 1 To 4
                With Cells(X + j - 1, Y + k - 1)
                    ' Set the cell properties if the cell is part of the block
                    If Arr(1 + j, 1 + k) = 1 Then
                        .Interior.Color = Nor  ' Set the interior color
                        
                        ' Set the top border properties
                        If Arr(0 + j, 1 + k) = 1 Then
                            .Borders(8).LineStyle = -4142  ' No border
                        Else
                            .Borders(8).Color = Bri  ' Bright color
                            .Borders(8).Weight = 4  ' Border weight
                        End If
                        
                        ' Set the bottom border properties
                        If Arr(1 + j, 0 + k) = 1 Then
                            .Borders(7).LineStyle = -4142  ' No border
                        Else
                            .Borders(7).Color = Bri  ' Bright color
                            .Borders(7).Weight = 4  ' Border weight
                        End If
                        
                        ' Set the left border properties
                        If Arr(2 + j, 1 + k) = 1 Then
                            .Borders(9).LineStyle = -4142  ' No border
                        Else
                            .Borders(9).Color = Dar  ' Dark color
                            .Borders(9).Weight = 4  ' Border weight
                        End If
                        
                        ' Set the right border properties
                        If Arr(1 + j, 2 + k) = 1 Then
                            .Borders(10).LineStyle = -4142  ' No border
                        Else
                            .Borders(10).Color = Dar  ' Dark color
                            .Borders(10).Weight = 4  ' Border weight
                        End If
                        
                        RowVal = RowVal + 1  ' Increment the row value
                    End If
                End With
            Next k
            If RowVal > 0 Then BloHei = BloHei + 1  ' Increment block height if row has value
        Next j
        
        X = X + BloHei + 2  ' Update the X position for the next block
    Next i
    
End Sub

Sub DisplayStatistics()
    ' Update Score
    With Cells(StaFie.X + 1, StaFie.Y + 7)
        If Sta.Sco <> .Value Then .Value = Sta.Sco  ' Update cell value if different
    End With
    
    ' Update Max Score
    If Sta.Sco > Sta.ScoMax Then Sta.ScoMax = Sta.Sco  ' Update max score if current score is higher
    With Cells(StaFie.X + 3, StaFie.Y + 7)
        If Sta.ScoMax <> .Value Then .Value = Sta.ScoMax  ' Update cell value if different
    End With

    ' Update Level
    With Cells(StaFie.X + 6, StaFie.Y + 7)
        If Sta.Lev <> .Value Then .Value = Sta.Lev  ' Update cell value if different
    End With

    ' Update Blocks
    With Cells(StaFie.X + 8, StaFie.Y + 7)
        If Sta.Blo <> .Value Then .Value = Sta.Blo  ' Update cell value if different
    End With

    ' Update Rows
    With Cells(StaFie.X + 10, StaFie.Y + 7)
        If Sta.Row <> .Value Then .Value = Sta.Row  ' Update cell value if different
    End With

    ' Update Quads
    With Cells(StaFie.X + 12, StaFie.Y + 7)
        If Sta.Qua <> .Value Then .Value = Sta.Qua  ' Update cell value if different
    End With

    ' Update Gapless
    With Cells(StaFie.X + 14, StaFie.Y + 7)
        If Sta.Gap <> .Value Then .Value = Sta.Gap  ' Update cell value if different
    End With
End Sub

Sub ClearMatrix(All As Byte)
    ' Loop through the game matrix
    For i = 4 To PlaFie.H + 3
        For j = 4 To PlaFie.W + 3
            ' Clear the matrix based on the value of All
            If All = 1 Then Mat(i, j) = 0  ' Clear the entire matrix if All is 1
            MatCop(i, j) = 0  ' Clear the matrix copy
        Next j
    Next i
End Sub





