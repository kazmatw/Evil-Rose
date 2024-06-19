Attribute VB_Name = "ModUtilities"
Sub SwitchToEnglish()
    ActivateKeyboardLayout HKL_ENGLISH, KLF_SETFORPROCESS
End Sub

Sub SwitchToChineseBopomofo()
    ActivateKeyboardLayout HKL_CHINESE_TRADITIONAL_PHONETIC, KLF_SETFORPROCESS
End Sub
Function ChangeBrightness(Col As Long, Per As Integer) As Long
    ' Declare variables for the RGB components of a color
    Dim R, G, B As Single
    
    ' Extract the Red, Green, and Blue components from the composite long integer color value
    R = Col Mod 256  ' Get the red component
    G = Col \ 256 Mod 256  ' Get the green component
    B = Col \ 65536 Mod 256  ' Get the blue component
    
    ' If brightness percentage is positive and any color component is at its maximum, adjust zero components slightly to prevent visual issues
    If Per > 0 Then
        If R = 255 Or G = 255 Or B = 255 Then
            If R = 0 Then R = 32  ' Increase red if initially zero
            If G = 0 Then G = 32  ' Increase green if initially zero
            If B = 0 Then B = 32  ' Increase blue if initially zero
        End If
    End If
    
    ' Adjust the RGB components based on the percentage (Per)
    R = R + Per * (R / 100)  ' Calculate new red value
    G = G + Per * (G / 100)  ' Calculate new green value
    B = B + Per * (B / 100)  ' Calculate new blue value
    
    ' Ensure RGB values remain within the 0-255 range
    If R < 0 Then R = 0
    If G < 0 Then G = 0
    If B < 0 Then B = 0
    If R > 255 Then R = 255
    If G > 255 Then G = 255
    If B > 255 Then B = 255
    
    ' Combine the adjusted RGB values into a single long integer and return it
    ChangeBrightness = RGB(Int(R), Int(G), Int(B))
End Function

Function CopyBloLibArrToCurBloArr(Blo As Byte)
    ' Loop through each element in the block's array (assuming 6x6 grid)
    For i = 1 To 6
        For j = 1 To 6
            ' Copy each block definition from the block library to the current block's array
            CurBlo.Arr(i, j) = BloLib(BloSet(CurBloSet).Blo(Blo)).Arr(i, j)
        Next j
    Next i
End Function

Function CopyBloLibArrToCurBloArr_2p(Blo_2p As Byte)
    ' Loop through each element in the block's array (assuming 6x6 grid)
    For i = 1 To 6
        For j = 1 To 6
            ' Copy each block definition from the block library to the current block's array
            CurBlo_2p.Arr(i, j) = BloLib(BloSet(CurBloSet).Blo(Blo_2p)).Arr(i, j)
        Next j
    Next i
End Function

Sub OpenGithub()
    ActiveWorkbook.FollowHyperlink Address:="https://github.com/kazmatw/Evil-Rose.git"

End Sub

Sub SetColor_GG()
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
    
    CurColSet = 16  'Change all Color to gray
    
    ' Loop through the game matrix dimensions
    For i = 4 To H + 3
        For j = 4 To W + 3
            ' Check if the current cell needs updating
            With Cells(X + i - 4, Y + j - 4)
                If Mat(i, j) <> 255 Then
                    Select Case Mat(i, j)
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
                            .Value = "X"  ' Placeholder value for testing
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
                End If
            End With
        Next j
    Next i
End Sub


