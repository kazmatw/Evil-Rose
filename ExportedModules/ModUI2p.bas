Attribute VB_Name = "ModUI2p"
Sub Init_GameSheetSize_2p()

    Set twoP = Worksheets("2p")
    ' Set all columns width to 4
    twoP.Cells.ColumnWidth = 4
    'Set all rows height to 20.1
    twoP.Cells.RowHeight = 20.1

End Sub

Sub CreateGameSheet_2p()

    ' Declare variables for colors, dimensions, and positions of the game and statistics fields
    Dim PFBBC, PFBNC, PFBDC, PFBBC_2p, PFBNC_2p, PFBDC_2p As Long ' Game field border colors
    Dim PFBC1, PFBC2, PFBC1_2p, PFBC2_2p As Long ' Game field background colors
    Dim PFH, PFW, PFH_2p, PFW_2p As Byte ' Game field height and width
    Dim PFX, PFY, PFX_2p, PFY_2p As Byte ' Game field position
    Dim SFBBC, SFBNC, SFBDC, SFBBC_2p, SFBNC_2p, SFBDC_2p As Long ' Statistics field border colors
    Dim SFBC1, SFBC2, SFBC1_2p, SFBC2_2p As Long ' Statistics field background colors
    Dim SFH, SFW, SFH_2p, SFW_2p As Byte ' Statistics field height and width
    Dim SFX, SFY, SFX_2p, SFY_2p As Byte ' Statistics field position
    
    ' Load game field properties from PlaFie
    With PlaFie
        PFBC1 = .BacCol1
        PFBC2 = .BacCol2
        PFBBC = .BorBCol
        PFBDC = .BorDCol
        PFBNC = .BorNCol
        PFH = .H '16
        PFW = .W '8
        PFX = .X '3
        PFY = .Y '3
    End With
    
    With PlaFie_2p
        PFBC1_2p = .BacCol1
        PFBC2_2p = .BacCol2
        PFBBC_2p = .BorBCol
        PFBDC_2p = .BorDCol
        PFBNC_2p = .BorNCol
        PFH_2p = .H '16
        PFW_2p = .W '8
        PFX_2p = .X '3
        PFY_2p = .Y '32
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
    
    With StaFie_2p
        SFBC1_2p = .BacCol1
        SFBC2_2p = .BacCol2
        SFBBC_2p = .BorBCol
        SFBDC_2p = .BorDCol
        SFBNC_2p = .BorNCol
        .H = PFH
        .Y = PFY_2p + PFW_2p + 3
        SFW_2p = .W
        SFH_2p = .H
        SFX_2p = .X
        SFY_2p = .Y
    End With
        
    ' Reset cells in the range A1:AF26
    With Range("A1:AX26")
        .Value = ""  ' Clear cell values
        .Borders.LineStyle = -4142  ' Remove cell borders
        .UnMerge  ' Unmerge any merged cells
    End With
    
    ' Draw game sheet background
    Range(Cells(PFX - 2, PFY - 2), Cells(PFX_2p + PFH_2p + 1, PFY_2p + PFW_2p + SFW_2p + 4)).Interior.Color = xlNone
    Range(Cells(PFX - 2, PFY - 2), Cells(PFX - 2, PFY_2p + PFW_2p + SFW_2p + 3)).Interior.Color = GamSheBC
    Range(Cells(PFX + PFH + 1, PFY - 2), Cells(PFX + PFH + 1, PFY_2p + PFW_2p + SFW_2p + 3)).Interior.Color = GamSheBC
    Range(Cells(PFX - 2, PFY - 2), Cells(PFX + PFH + 1, PFY - 2)).Interior.Color = GamSheBC
    Range(Cells(PFX - 2, PFY + PFW + 1), Cells(PFX + PFH + 1, PFY + PFW + 1)).Interior.Color = GamSheBC
    Range(Cells(PFX - 2, PFY + PFW + SFW_2p + 4), Cells(PFX + PFH + 1, PFY + PFW + SFW_2p + 4)).Interior.Color = GamSheBC
    Range(Cells(PFX_2p - 2, PFY_2p - 2), Cells(PFX_2p + PFH_2p + 1, PFY_2p - 2)).Interior.Color = GamSheBC
    Range(Cells(PFX_2p - 2, PFY_2p + PFW_2p + 1), Cells(PFX_2p + PFH_2p + 1, PFY_2p + PFW_2p + 1)).Interior.Color = GamSheBC
    Range(Cells(PFX_2p - 2, PFY_2p + PFW_2p + SFW_2p + 4), Cells(PFX_2p + PFH_2p + 1, PFY_2p + PFW_2p + SFW_2p + 4)).Interior.Color = GamSheBC
    Range(Cells(PFX - 2, PFY + PFW + SFW_2p + 5), Cells(PFX + PFH + 1, PFY_2p - 3)).Interior.Color = xlNone
   
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
    
    With Range(Cells(PFX_2p - 1, PFY_2p - 1), Cells(PFX_2p + PFH_2p, PFY_2p + PFW_2p))
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
    Range(Cells(PFX_2p - 1, PFY_2p - 1), Cells(PFX_2p + PFH_2p, PFY_2p + PFW_2p)).Interior.Color = PFBNC
    
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
    
    With Range(Cells(PFX_2p, PFY_2p), Cells(PFX_2p + PFH_2p - 1, PFY_2p + PFW_2p - 1))
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
    
    For i = SFX To SFX + SFH - 1
        Range(Cells(i, SFY), Cells(i, SFY + 5)).Merge ' Merge cells for score display
        Range(Cells(i, SFY_2p), Cells(i, SFY_2p + 5)).Merge
    Next i
    
    With Range(Cells(SFX - 1, SFY - 1), Cells(SFX + SFH, SFY + SFW_2p))
        .Borders(8).Color = SFBBC  ' Top border
        .Borders(8).Weight = 4
        .Borders(7).Color = SFBBC  ' Bottom border
        .Borders(7).Weight = 4
        .Borders(9).Color = SFBDC  ' Left border
        .Borders(9).Weight = 4
        .Borders(10).Color = SFBDC  ' Right border
        .Borders(10).Weight = 4
    End With
    
    With Range(Cells(SFX_2p - 1, SFY_2p - 1), Cells(SFX_2p + SFH_2p, SFY_2p + SFW_2p))
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
    Range(Cells(SFX - 1, SFY - 1), Cells(SFX - 1, SFY + SFW_2p)).Interior.Color = SFBNC
    Range(Cells(SFX + SFH, SFY - 1), Cells(SFX + SFH, SFY + SFW_2p)).Interior.Color = SFBNC
    Range(Cells(SFX - 1, SFY - 1), Cells(SFX + SFH, SFY - 1)).Interior.Color = SFBNC
    Range(Cells(SFX - 1, SFY + 6), Cells(SFX + SFH, SFY + 6)).Interior.Color = SFBNC
    
    Range(Cells(SFX_2p - 1, SFY_2p - 1), Cells(SFX_2p - 1, SFY_2p + SFW_2p)).Interior.Color = SFBNC
    Range(Cells(SFX_2p + SFH_2p, SFY_2p - 1), Cells(SFX_2p + SFH_2p, SFY_2p + SFW_2p)).Interior.Color = SFBNC
    Range(Cells(SFX_2p - 1, SFY_2p - 1), Cells(SFX_2p + SFH_2p, SFY_2p - 1)).Interior.Color = SFBNC
    Range(Cells(SFX_2p - 1, SFY_2p + 6), Cells(SFX_2p + SFH_2p, SFY_2p + 6)).Interior.Color = SFBNC
    ' Set interior properties for the score display cells
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
    
    With Range(Cells(SFX_2p, SFY_2p), Cells(SFX_2p + SFH_2p - 1, SFY_2p + 5))
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


    ' Configure score and max score cells
    For i = SFX To SFX + 2 Step 2
        With Cells(i, SFY)
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
        With Cells(i + 1, SFY)
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
    
    For i = SFX_2p To SFX_2p + 2 Step 2
        With Cells(i, SFY_2p)
            .Font.Color = &H884444
            .Font.Bold = True
            .Font.Italic = True
            .Font.Name = "Arial"
            .Font.Size = 18
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            If i = SFX_2p Then
                .Value = "SCORE"
            Else
                .Value = "MAX SCORE"
            End If
        End With
        With Cells(i + 1, SFY_2p)
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
        With Cells(i, SFY)
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
        With Cells(i + 1, SFY)
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
        With Cells(i, SFY_2p)
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
        With Cells(i + 1, SFY_2p)
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

Sub DrawPlayingField_2p(Mode_2p As Byte)
    ' Declare static variable to hold the previous level
    Static PreLev_2p As Byte
    ' Declare variables for coordinates and dimensions
    Dim X, Y As Byte
    Dim H, W As Byte
    
    ' Load game field properties from PlaFie
    With PlaFie_2p
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
            If Mat_2p(i, j) <> MatCop_2p(i, j) Or Mode_2p = 1 And Sta_2p.Lev <> PreLev_2p Or Mode_2p = 2 Then
                ' Update the cell in the game field
                With Cells(X + i - 4, Y + j - 4)
                    Select Case Mat_2p(i, j)
                        ' Handle the case for current block (value 255)
                        Case 255
                            .Value = ""
                            .Interior.Color = CurBlo_2p.NorCol  ' Set the normal color
                            ' Set the top border color and weight
                            If Mat_2p(i - 1, j) < 255 Then
                                .Borders(8).Color = CurBlo_2p.BriCol
                                .Borders(8).Weight = 4
                            Else
                                .Borders(8).LineStyle = -4142
                            End If
                            ' Set the bottom border color and weight
                            If Mat_2p(i, j - 1) < 255 Then
                                .Borders(7).Color = CurBlo_2p.BriCol
                                .Borders(7).Weight = 4
                            Else
                                .Borders(7).LineStyle = -4142
                            End If
                            ' Set the left border color and weight
                            If Mat_2p(i + 1, j) < 255 Then
                                .Borders(9).Color = CurBlo_2p.DarCol
                                .Borders(9).Weight = 4
                            Else
                                .Borders(9).LineStyle = -4142
                            End If
                            ' Set the right border color and weight
                            If Mat_2p(i, j + 1) < 255 Then
                                .Borders(10).Color = CurBlo_2p.DarCol
                                .Borders(10).Weight = 4
                            Else
                                .Borders(10).LineStyle = -4142
                            End If
                        ' Handle the case for normal block (value > 0)
                        Case Is > 0
                            .Value = ""
                            .Interior.Color = ColLib(ColSet(CurColSet, Mat_2p(i, j))).Nor  ' Set the normal color
                            ' Set the top border color and weight
                            If i > 4 Then
                                If Mat_2p(i - 1, j) = 0 Then
                                    .Borders(8).Color = ColLib(ColSet(CurColSet, Mat_2p(i, j))).Bri
                                    .Borders(8).Weight = 2
                                ElseIf Mat_2p(i - 1, j) <> Mat_2p(i, j) And Mat_2p(i - 1, j) < 255 Then
                                    .Borders(8).Color = &H444444
                                    .Borders(8).Weight = 2
                                Else
                                    .Borders(8).LineStyle = -4142
                                End If
                            Else
                                .Borders(8).Color = PlaFie_2p.BorDCol
                                .Borders(8).Weight = 4
                            End If
                            ' Set the bottom border color and weight
                            If j > 4 Then
                                If Mat_2p(i, j - 1) = 0 Then
                                    .Borders(7).Color = ColLib(ColSet(CurColSet, Mat_2p(i, j))).Bri
                                    .Borders(7).Weight = 2
                                ElseIf Mat_2p(i, j - 1) <> Mat_2p(i, j) And Mat_2p(i, j - 1) < 255 Then
                                    .Borders(7).Color = &H444444
                                    .Borders(7).Weight = 2
                                Else
                                    .Borders(7).LineStyle = -4142
                                End If
                            Else
                                .Borders(7).Color = PlaFie_2p.BorDCol
                                .Borders(7).Weight = 4
                            End If
                            ' Set the left border color and weight
                            If i < H + 3 Then
                                If Mat_2p(i + 1, j) = 0 Then
                                    .Borders(9).Color = ColLib(ColSet(CurColSet, Mat_2p(i, j))).Dar
                                    .Borders(9).Weight = 2
                                ElseIf Mat_2p(i + 1, j) <> Mat_2p(i, j) And Mat_2p(i + 1, j) < 255 Then
                                    .Borders(9).Color = &H444444
                                    .Borders(9).Weight = 2
                                Else
                                    .Borders(9).LineStyle = -4142
                                End If
                            Else
                                .Borders(9).Color = PlaFie_2p.BorBCol
                                .Borders(9).Weight = 4
                            End If
                            ' Set the right border color and weight
                            If j < W + 3 Then
                                If Mat_2p(i, j + 1) = 0 Then
                                    .Borders(10).Color = ColLib(ColSet(CurColSet, Mat_2p(i, j))).Dar
                                    .Borders(10).Weight = 2
                                ElseIf Mat_2p(i, j + 1) <> Mat_2p(i, j) And Mat_2p(i, j + 1) < 255 Then
                                    .Borders(10).Color = &H444444
                                    .Borders(10).Weight = 2
                                Else
                                    .Borders(10).LineStyle = -4142
                                End If
                            Else
                                .Borders(10).Color = PlaFie_2p.BorBCol
                                .Borders(10).Weight = 4
                            End If
                        ' Handle the case for empty cell (value = 0)
                        Case Else
                            .Interior.Color = PlaFie_2p.BacCol1  ' Set the background color
                            .Value = "X"  ' Placeholder value for testing
                            ' Set the top border color and weight
                            If i = 4 Then
                                .Borders(8).Color = PlaFie_2p.BorDCol
                                .Borders(8).Weight = 4
                            ElseIf Mat_2p(i - 1, j) = 255 Then
                                .Borders(8).Color = CurBlo_2p.DarCol
                                .Borders(8).Weight = 4
                            ElseIf Mat_2p(i - 1, j) > 0 Then
                                .Borders(8).Color = ColLib(ColSet(CurColSet, Mat_2p(i - 1, j))).Dar
                                .Borders(8).Weight = 2
                            Else
                                .Borders(8).LineStyle = -4142
                            End If
                            ' Set the bottom border color and weight
                            If j = 4 Then
                                .Borders(7).Color = PlaFie_2p.BorDCol
                                .Borders(7).Weight = 4
                            ElseIf Mat_2p(i, j - 1) = 255 Then
                                .Borders(7).Color = CurBlo.DarCol
                                .Borders(7).Weight = 4
                            ElseIf Mat_2p(i, j - 1) > 0 Then
                                .Borders(7).Color = ColLib(ColSet(CurColSet, Mat_2p(i, j - 1))).Dar
                                .Borders(7).Weight = 2
                            Else
                                .Borders(7).LineStyle = -4142
                            End If
                            ' Set the left border color and weight
                            If i = H + 3 Then
                                .Borders(9).Color = PlaFie_2p.BorBCol
                                .Borders(9).Weight = 4
                            ElseIf Mat_2p(i + 1, j) = 255 Then
                                .Borders(9).Color = CurBlo_2p.BriCol
                                .Borders(9).Weight = 4
                            ElseIf Mat_2p(i + 1, j) > 0 Then
                                .Borders(9).Color = ColLib(ColSet(CurColSet, Mat_2p(i + 1, j))).Bri
                                .Borders(9).Weight = 2
                            Else
                                .Borders(9).LineStyle = -4142
                            End If
                            ' Set the right border color and weight
                            If j = W + 3 Then
                                .Borders(10).Color = PlaFie_2p.BorBCol
                                .Borders(10).Weight = 4
                            ElseIf Mat_2p(i, j + 1) = 255 Then
                                .Borders(10).Color = CurBlo_2p.BriCol
                                .Borders(10).Weight = 4
                            ElseIf Mat_2p(i, j + 1) > 0 Then
                                .Borders(10).Color = ColLib(ColSet(CurColSet, Mat_2p(i, j + 1))).Bri
                                .Borders(10).Weight = 2
                            Else
                                .Borders(10).LineStyle = -4142
                            End If
                    End Select
                End With
                ' Update the matrix copy to reflect the changes
                MatCop_2p(i, j) = Mat_2p(i, j)
            End If
        Next j
    Next i
    
    ' Store the current level as the previous level for the next update
    PreLev_2p = Sta_2p.Lev
End Sub

Sub DrawDeletedRows_2p()
    ' Declare variables for coordinates of the game field
    Dim X, Y As Byte
    
    ' Load game field position from PlaFie
    X = PlaFie_2p.X  ' X coordinate of the game field
    Y = PlaFie_2p.Y  ' Y coordinate of the game field

    ' Loop through each row in the game matrix
    For i = 4 To PlaFie_2p.H + 3
        ' Check if the current row has changed (indicating it was deleted)
        If Mat_2p(i, 4) <> MatCop_2p(i, 4) Then
            ' Update the cells in the row to reflect the deletion
            With Range(Cells(X + i - 4, Y), Cells(X + i - 4, Y + PlaFie_2p.W - 1))
                .Borders.LineStyle = -4142  ' Remove all borders
                .Borders(7).Color = PlaFie_2p.BorDCol  ' Set bottom border color
                .Borders(7).Weight = 4  ' Set bottom border weight
                .Borders(10).Color = PlaFie_2p.BorBCol  ' Set right border color
                .Borders(10).Weight = 4  ' Set right border weight
                ' Set top border for the first row
                If i = 4 Then
                    .Borders(8).Color = PlaFie_2p.BorDCol  ' Set top border color
                    .Borders(8).Weight = 4  ' Set top border weight
                End If
                ' Set bottom border for the last row
                If i = PlaFie.H + 3 Then
                    .Borders(9).Color = PlaFie_2p.BorBCol  ' Set bottom border color
                    .Borders(9).Weight = 4  ' Set bottom border weight
                End If
                .Interior.Color = PlaFie_2p.BacCol1  ' Set interior color
                .Value = "X"  ' Set placeholder value for testing
            End With
        End If
    Next i
    
    ' Update the game state to indicate rows have been deleted
    GamSta = 3
End Sub

Sub DisplayStatistics_2p()
    ' Update Score
    With Cells(StaFie.X + 1, StaFie.Y)
        If Sta.Sco <> .Value Then .Value = Sta.Sco  ' Update cell value if different
    End With
    
    ' Update Max Score
    If Sta.Sco > Sta.ScoMax Then Sta.ScoMax = Sta.Sco  ' Update max score if current score is higher
    With Cells(StaFie.X + 3, StaFie.Y)
        If Sta.ScoMax <> .Value Then .Value = Sta.ScoMax  ' Update cell value if different
    End With

    ' Update Level
    With Cells(StaFie.X + 6, StaFie.Y)
        If Sta.Lev <> .Value Then .Value = Sta.Lev  ' Update cell value if different
    End With

    ' Update Blocks
    With Cells(StaFie.X + 8, StaFie.Y)
        If Sta.Blo <> .Value Then .Value = Sta.Blo  ' Update cell value if different
    End With

    ' Update Rows
    With Cells(StaFie.X + 10, StaFie.Y)
        If Sta.Row <> .Value Then .Value = Sta.Row  ' Update cell value if different
    End With

    ' Update Quads
    With Cells(StaFie.X + 12, StaFie.Y)
        If Sta.Qua <> .Value Then .Value = Sta.Qua  ' Update cell value if different
    End With

    ' Update Gapless
    With Cells(StaFie.X + 14, StaFie.Y)
        If Sta.Gap <> .Value Then .Value = Sta.Gap  ' Update cell value if different
    End With
'--------------------------------------------------------------------------------------------------------------------
    ' Update Score
    With Cells(StaFie_2p.X + 1, StaFie_2p.Y)
        If Sta_2p.Sco <> .Value Then .Value = Sta_2p.Sco  ' Update cell value if different
    End With
    
    ' Update Max Score
    If Sta_2p.Sco > Sta_2p.ScoMax Then Sta_2p.ScoMax = Sta_2p.Sco  ' Update max score if current score is higher
    With Cells(StaFie_2p.X + 3, StaFie_2p.Y)
        If Sta_2p.ScoMax <> .Value Then .Value = Sta_2p.ScoMax  ' Update cell value if different
    End With

    ' Update Level
    With Cells(StaFie_2p.X + 6, StaFie_2p.Y)
        If Sta_2p.Lev <> .Value Then .Value = Sta_2p.Lev  ' Update cell value if different
    End With

    ' Update Blocks
    With Cells(StaFie_2p.X + 8, StaFie_2p.Y)
        If Sta_2p.Blo <> .Value Then .Value = Sta_2p.Blo  ' Update cell value if different
    End With

    ' Update Rows
    With Cells(StaFie_2p.X + 10, StaFie_2p.Y)
        If Sta_2p.Row <> .Value Then .Value = Sta_2p.Row  ' Update cell value if different
    End With

    ' Update Quads
    With Cells(StaFie_2p.X + 12, StaFie_2p.Y)
        If Sta_2p.Qua <> .Value Then .Value = Sta_2p.Qua  ' Update cell value if different
    End With

    ' Update Gapless
    With Cells(StaFie_2p.X + 14, StaFie_2p.Y)
        If Sta_2p.Gap <> .Value Then .Value = Sta_2p.Gap  ' Update cell value if different
    End With
End Sub

Sub ClearMatrix_2p(All As Byte)
    ' Loop through the game matrix
    For i = 4 To PlaFie_2p.H + 3
        For j = 4 To PlaFie_2p.W + 3
            ' Clear the matrix based on the value of All
            If All = 1 Then Mat_2p(i, j) = 0  ' Clear the entire matrix if All is 1
            MatCop_2p(i, j) = 0  ' Clear the matrix copy
        Next j
    Next i
End Sub
