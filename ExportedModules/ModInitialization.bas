Attribute VB_Name = "ModInitialization"
Sub SetInitialValues()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Game") ' Change "Game" to your worksheet name
    
    ' Set all columns width to 4
    ws.Cells.ColumnWidth = 4
    
    ' Set all rows height to 20.1
    ws.Cells.RowHeight = 20.1

    'Blocks
    
    Dim DatStr As String
    
    DatStr = DatStr + "300000" + _
                      "010000" + _
                      "011100" + _
                      "000000" + _
                      "000000" + _
                      "000000"

    DatStr = DatStr + "300000" + _
                      "000100" + _
                      "011100" + _
                      "000000" + _
                      "000000" + _
                      "000030"

    DatStr = DatStr + "300000" + _
                      "011000" + _
                      "001100" + _
                      "000000" + _
                      "000000" + _
                      "000000"

    DatStr = DatStr + "300000" + _
                      "001100" + _
                      "011000" + _
                      "000000" + _
                      "000000" + _
                      "000000"

    DatStr = DatStr + "300000" + _
                      "001000" + _
                      "011100" + _
                      "000000" + _
                      "000000" + _
                      "000000"

    DatStr = DatStr + "200000" + _
                      "011000" + _
                      "011000" + _
                      "000000" + _
                      "000000" + _
                      "000000"

    DatStr = DatStr + "400000" + _
                      "000000" + _
                      "011110" + _
                      "000000" + _
                      "000000" + _
                      "000000"

    Dim l As Integer
    Dim DatVal As Byte
    l = 1
    For i = 1 To 7
        For j = 1 To 6
            For k = 1 To 6
                DatVal = Val(Mid(DatStr, l, 1))
                If j + k = 2 Then
                    BloLib(i).Siz = DatVal
                Else
                    BloLib(i).Arr(j, k) = DatVal
                End If
                l = l + 1
            Next k
        Next j
    Next i
    
    ' Ensure CurBloSet and CurColSet are within bounds
    CurBloSet = 1
    CurColSet = 1
    
    ' Blocks and block sets
    ReDim BloSet(1 To 1)
    With BloSet(1)
        ReDim .Blo(1 To 7)
        .Blo(1) = 1
        .Blo(2) = 2
        .Blo(3) = 3
        .Blo(4) = 4
        .Blo(5) = 5
        .Blo(6) = 6
        .Blo(7) = 7
    End With

    'Color Library

    ColLib(1).Nor = RGB(255, 0, 0)          'Red
    ColLib(2).Nor = RGB(0, 255, 0)          'Green
    ColLib(3).Nor = RGB(0, 0, 255)          'Blue
    ColLib(4).Nor = RGB(255, 255, 0)        'Yellow
    ColLib(5).Nor = RGB(255, 0, 255)        'Magenta
    ColLib(6).Nor = RGB(0, 255, 255)        'Cyan
    ColLib(7).Nor = RGB(255, 128, 0)        'Orange
    
    ColLib(8).Nor = RGB(250, 235, 215)      'Antique White
    ColLib(9).Nor = RGB(255, 239, 213)      'Papaya Whip
    ColLib(10).Nor = RGB(255, 235, 205)     'Blanched Almond
    ColLib(11).Nor = RGB(255, 228, 196)     'Bisque
    ColLib(12).Nor = RGB(255, 218, 185)     'Peach Puff
    ColLib(13).Nor = RGB(255, 222, 173)     'Navajo White
    ColLib(14).Nor = RGB(255, 228, 181)     'Moccasin
    
    ColLib(15).Nor = RGB(119, 136, 153)     'Light Slate Gray
    ColLib(16).Nor = RGB(190, 190, 190)     'Grey
    ColLib(17).Nor = RGB(211, 211, 211)     'Light Gray
    ColLib(18).Nor = RGB(25, 25, 112)       'Midnight Blue
    ColLib(19).Nor = RGB(100, 149, 237)     'Cornflower Blue
    ColLib(20).Nor = RGB(72, 61, 139)       'Dark Slate Blue
    ColLib(21).Nor = RGB(106, 90, 205)      'Slate Blue
    
    ColLib(22).Nor = RGB(30, 144, 255)      'Dodger Blue
    ColLib(23).Nor = RGB(0, 191, 255)       'Deep Sky Blue
    ColLib(24).Nor = RGB(135, 206, 235)     'Sky Blue
    ColLib(25).Nor = RGB(176, 196, 222)     'Light Steel Blue
    ColLib(26).Nor = RGB(0, 206, 209)       'Dark Turquoise
    ColLib(27).Nor = RGB(85, 107, 47)       'Dark Olive Green
    ColLib(28).Nor = RGB(60, 179, 113)      'Medium Sea Green
    
    ColLib(29).Nor = RGB(32, 178, 170)      'Light Sea Green
    ColLib(30).Nor = RGB(152, 251, 152)     'Pale Green
    ColLib(31).Nor = RGB(0, 255, 127)       'Spring Green
    ColLib(32).Nor = RGB(0, 250, 154)       'Med Spring Green
    ColLib(33).Nor = RGB(173, 255, 47)      'Green Yellow
    ColLib(34).Nor = RGB(107, 142, 35)      'Olive Drab
    ColLib(35).Nor = RGB(189, 183, 107)     'Dark Khaki
    
    ColLib(36).Nor = RGB(238, 232, 170)     'Pale Goldenrod
    ColLib(37).Nor = RGB(250, 250, 210)     'Lt Goldenrod Yello
    ColLib(38).Nor = RGB(255, 215, 0)       'Gold
    ColLib(39).Nor = RGB(218, 165, 32)      'Golden Rod
    ColLib(40).Nor = RGB(188, 143, 143)     'Rosy Brown
    ColLib(41).Nor = RGB(205, 92, 92)       'Indian Red
    ColLib(42).Nor = RGB(205, 133, 63)      'Peru
    
    ColLib(43).Nor = RGB(245, 222, 179)     'Wheat
    ColLib(44).Nor = RGB(244, 164, 96)      'Sandy Brown
    ColLib(45).Nor = RGB(210, 180, 140)     'Tan
    ColLib(46).Nor = RGB(178, 34, 34)       'Firebrick
    ColLib(47).Nor = RGB(250, 128, 114)     'Salmon
    ColLib(48).Nor = RGB(255, 160, 122)     'Light Salmon
    
    ColLib(49).Nor = RGB(255, 99, 71)       'Tomato
    ColLib(50).Nor = RGB(255, 69, 0)        'Orange Red
    ColLib(51).Nor = RGB(255, 105, 180)     'Hot Pink
    ColLib(52).Nor = RGB(255, 20, 147)      'Deep Pink
    ColLib(53).Nor = RGB(255, 182, 193)     'Light Pink
    ColLib(54).Nor = RGB(219, 112, 147)     'Pale Violet Red
    ColLib(55).Nor = RGB(176, 48, 96)       'Maroon
    
    ColLib(56).Nor = RGB(218, 112, 214)     'Orchid
    ColLib(57).Nor = RGB(186, 85, 211)      'Medium Orchid
    ColLib(58).Nor = RGB(148, 0, 211)       'Dark Violet
    ColLib(59).Nor = RGB(138, 43, 226)      'Blue Violet
    ColLib(60).Nor = RGB(216, 191, 216)     'Thistle
    ColLib(61).Nor = RGB(238, 233, 233)     'Snow
    ColLib(62).Nor = RGB(205, 197, 191)     'Seashell
    
    ColLib(63).Nor = RGB(255, 218, 185)     'Peach Puff 1
    ColLib(64).Nor = RGB(238, 203, 173)     'Peach Puff 2
    ColLib(65).Nor = RGB(255, 222, 173)     'Navajo White 1
    ColLib(66).Nor = RGB(205, 179, 139)     'Navajo White 2
    ColLib(67).Nor = RGB(238, 233, 191)     'Lemon Chiffon 1
    ColLib(68).Nor = RGB(205, 201, 165)     'Lemon Chiffon 2
    ColLib(69).Nor = RGB(139, 136, 120)     'Cornsilk
    
    ColLib(70).Nor = RGB(205, 193, 197)     'Lavender Blush
    ColLib(71).Nor = RGB(238, 213, 210)     'Misty Rose
    ColLib(72).Nor = RGB(193, 205, 205)     'Azure
    ColLib(73).Nor = RGB(72, 118, 255)      'Royal Blue
    ColLib(74).Nor = RGB(176, 226, 255)     'Light Sky Blue
    ColLib(75).Nor = RGB(150, 205, 205)     'Pale Turquoise
    ColLib(76).Nor = RGB(142, 229, 238)     'Cadet Blue
    
    ColLib(77).Nor = RGB(127, 255, 212)     'Aquamarine 1
    ColLib(78).Nor = RGB(102, 205, 170)     'Aquamarine 2
    ColLib(79).Nor = RGB(180, 238, 180)     'Dark Sea Green
    ColLib(80).Nor = RGB(144, 238, 144)     'Pale Green
    ColLib(81).Nor = RGB(0, 205, 0)         'Green 1
    ColLib(82).Nor = RGB(0, 139, 0)         'Green 2
    ColLib(83).Nor = RGB(188, 238, 104)     'Dark Olive Green2
    
    ColLib(84).Nor = RGB(255, 236, 139)     'Light Goldenrod 1
    ColLib(85).Nor = RGB(205, 190, 112)     'Light Goldenrod 2
    ColLib(86).Nor = RGB(205, 205, 180)     'Light Yellow
    ColLib(87).Nor = RGB(238, 238, 0)       'Yellow 1
    ColLib(88).Nor = RGB(205, 205, 0)       'Yellow 2
    ColLib(89).Nor = RGB(139, 139, 0)       'Yellow 3
    ColLib(90).Nor = RGB(238, 201, 0)       'Gold 1
    
    ColLib(91).Nor = RGB(255, 193, 37)      'Goldenrod 1
    ColLib(92).Nor = RGB(238, 180, 34)      'Goldenrod 2
    ColLib(93).Nor = RGB(205, 155, 29)      'Goldenrod 3
    ColLib(94).Nor = RGB(139, 105, 20)      'Goldenrod 4
    ColLib(95).Nor = RGB(255, 185, 15)      'Dark Goldenrod 1
    ColLib(96).Nor = RGB(238, 173, 14)      'Dark Goldenrod 2
    ColLib(97).Nor = RGB(205, 149, 12)      'Dark Goldenrod 3
    
    ColLib(98).Nor = RGB(255, 193, 193)     'Rosy Brown 1
    ColLib(99).Nor = RGB(205, 155, 155)     'Rosy Brown 2
    ColLib(100).Nor = RGB(255, 106, 106)    'Indian Red 1
    ColLib(101).Nor = RGB(205, 85, 85)      'Indian Red 2
    ColLib(102).Nor = RGB(255, 130, 71)     'Sienna 1
    ColLib(103).Nor = RGB(205, 104, 57)     'Sienna 2
    ColLib(104).Nor = RGB(238, 197, 145)    'Burlywood
    
    ColLib(105).Nor = RGB(0, 0, 139)        'Dark Blue
    ColLib(106).Nor = RGB(0, 139, 139)      'Dark Cyan
    ColLib(107).Nor = RGB(139, 0, 139)      'Dark Magenta
    ColLib(108).Nor = RGB(139, 0, 0)        'Dark Red
    ColLib(109).Nor = RGB(144, 238, 144)    'Light Green
    ColLib(110).Nor = RGB(161, 130, 103)    'Gold Brown
    ColLib(111).Nor = RGB(85, 88, 90)       'Platinum
    
    For i = 1 To 112
        ColLib(i).Bri = ChangeBrightness(ColLib(i).Nor, 240)
        ColLib(i).Dar = ChangeBrightness(ColLib(i).Nor, -40)
    Next i
    
    'Color Sets
    
    ReDim ColSet(15, 7)
    
        'Set 1
        
        ColSet(1, 1) = 2
        ColSet(1, 2) = 2
        ColSet(1, 3) = 5
        ColSet(1, 4) = 5
        ColSet(1, 5) = 3
        ColSet(1, 6) = 4
        ColSet(1, 7) = 1
        
        'Set 2
        
        ColSet(2, 1) = 50
        ColSet(2, 2) = 51
        ColSet(2, 3) = 52
        ColSet(2, 4) = 53
        ColSet(2, 5) = 54
        ColSet(2, 6) = 55
        ColSet(2, 7) = 56
        
        'Set 3
        
        ColSet(3, 1) = 15
        ColSet(3, 2) = 16
        ColSet(3, 3) = 17
        ColSet(3, 4) = 18
        ColSet(3, 5) = 19
        ColSet(3, 6) = 20
        ColSet(3, 7) = 21
        
        'Set 4
        
        ColSet(4, 1) = 22
        ColSet(4, 2) = 23
        ColSet(4, 3) = 24
        ColSet(4, 4) = 25
        ColSet(4, 5) = 26
        ColSet(4, 6) = 27
        ColSet(4, 7) = 28
        
        'Set 5
        
        ColSet(5, 1) = 29
        ColSet(5, 2) = 30
        ColSet(5, 3) = 31
        ColSet(5, 4) = 32
        ColSet(5, 5) = 33
        ColSet(5, 6) = 34
        ColSet(5, 7) = 35
        
        'Set 6
        
        ColSet(6, 1) = 36
        ColSet(6, 2) = 37
        ColSet(6, 3) = 38
        ColSet(6, 4) = 39
        ColSet(6, 5) = 40
        ColSet(6, 6) = 41
        ColSet(6, 7) = 42
        
        'Set 7
        
        ColSet(7, 1) = 43
        ColSet(7, 2) = 44
        ColSet(7, 3) = 45
        ColSet(7, 4) = 46
        ColSet(7, 5) = 47
        ColSet(7, 6) = 48
        ColSet(7, 7) = 49
        
        'Set 8
        
        ColSet(8, 1) = 8
        ColSet(8, 2) = 9
        ColSet(8, 3) = 10
        ColSet(8, 4) = 11
        ColSet(8, 5) = 12
        ColSet(8, 6) = 13
        ColSet(8, 7) = 14
        
        'Set 9
        
        ColSet(9, 1) = 57
        ColSet(9, 2) = 58
        ColSet(9, 3) = 59
        ColSet(9, 4) = 60
        ColSet(9, 5) = 61
        ColSet(9, 6) = 62
        ColSet(9, 7) = 63
        
        'Set 10
        
        ColSet(10, 1) = 64
        ColSet(10, 2) = 65
        ColSet(10, 3) = 66
        ColSet(10, 4) = 67
        ColSet(10, 5) = 68
        ColSet(10, 6) = 69
        ColSet(10, 7) = 70
        
        'Set 11
        
        ColSet(11, 1) = 71
        ColSet(11, 2) = 72
        ColSet(11, 3) = 73
        ColSet(11, 4) = 74
        ColSet(11, 5) = 75
        ColSet(11, 6) = 76
        ColSet(11, 7) = 77
         
        'Set 12
        
        ColSet(12, 1) = 78
        ColSet(12, 2) = 79
        ColSet(12, 3) = 80
        ColSet(12, 4) = 81
        ColSet(12, 5) = 82
        ColSet(12, 6) = 83
        ColSet(12, 7) = 84
        
        'Set 13
        
        ColSet(13, 1) = 85
        ColSet(13, 2) = 86
        ColSet(13, 3) = 87
        ColSet(13, 4) = 88
        ColSet(13, 5) = 89
        ColSet(13, 6) = 90
        ColSet(13, 7) = 91
        
        'Set 14
        
        ColSet(14, 1) = 92
        ColSet(14, 2) = 93
        ColSet(14, 3) = 94
        ColSet(14, 4) = 95
        ColSet(14, 5) = 96
        ColSet(14, 6) = 97
        ColSet(14, 7) = 98
        
        'Set 15
        
        ColSet(15, 1) = 99
        ColSet(15, 2) = 100
        ColSet(15, 3) = 101
        ColSet(15, 4) = 102
        ColSet(15, 5) = 103
        ColSet(15, 6) = 104
        ColSet(15, 7) = 105
    
    'Playing Field
    
    With PlaFie
        .BacCol1 = RGB(32, 32, 32)
        .BacCol2 = RGB(48, 48, 48)
        .BorBCol = RGB(224, 224, 224)
        .BorDCol = RGB(8, 8, 8)
        .BorNCol = RGB(128, 128, 128)
        .X = 3
        .Y = 3
    End With
    
    'Statistics Field
    
    With StaFie
        .BacCol1 = PlaFie.BacCol1
        .BacCol2 = PlaFie.BacCol2
        .BorBCol = PlaFie.BorBCol
        .BorDCol = PlaFie.BorDCol
        .BorNCol = PlaFie.BorNCol
        .W = 13
        .X = PlaFie.X
    End With
    BloPre = 1
    
    'Game Sheet Background Color
    
    GamSheBC = RGB(192, 192, 192)
    
    'Execution Threshold
    
    Tim.ExeThrDef = 5

End Sub

Sub InitializeGame()
    ' Create Playing Field Matrix
    PlaFie.H = 16
    PlaFie.W = 8
    ReDim Mat(PlaFie.H + 6, PlaFie.W + 6)
    ReDim MatCop(PlaFie.H + 6, PlaFie.W + 6)
    For i = 1 To PlaFie.H + 6
        For j = 1 To PlaFie.W + 6
            Mat(i, j) = 1
            MatCop(i, j) = 1
        Next j
    Next i

    ' Set current block and color set indices to valid values
    CurColSet = 1
    CurBloSet = 1
    
    ' Debug: Print current block set
    Debug.Print "CurBloSet: "; CurBloSet
    Debug.Print "UBound(BloSet): "; UBound(BloSet)
End Sub



Sub NewGame()

    With Sta
        .Blo = 0
        .Gap = 1
        .GapSum = 0
        .Lev = 1
        .LevPro = 0
        .Row = 0
        .Sco = 0
        .Qua = 0
    End With
    GamSta = 1
   
    'Initialize Statistics Display
    
    Call DisplayStatistics
    
    Tim.LevTim = 16
    Tim.ExeThr = Tim.LevTim
    Call ClearMatrix(1)
    Randomize
    Call GenerateBlocks(4)
    Call AssignKeys
    Call StartTimer
    
End Sub

