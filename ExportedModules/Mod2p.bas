Attribute VB_Name = "Mod2p"
Sub CreateGameSheet_2p()
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
        SFH = .H
        SFW = .W
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
    Range(Cells(PFX - 2, PFY - 2), Cells(PFX - 2, PFY + PFW + 1)).Interior.Color = GamSheBC  '(1,1)~(1,12)
    Range(Cells(PFX + PFH + 1, PFY - 2), Cells(PFX + PFH + 1, PFY + PFW + 1)).Interior.Color = GamSheBC ' (20,1)~(20,12)
    Range(Cells(PFX - 2, PFY - 2), Cells(PFX + PFH + 1, PFY - 2)).Interior.Color = GamSheBC '(1,1)~(20,1)
    Range(Cells(PFX - 2, PFY + PFW + 1), Cells(PFX + PFH + 1, PFY + PFW + 1)).Interior.Color = GamSheBC '(1,12)~(20,12)
    
    '2p
    Range(Cells(PFX - 2, PFY + 18), Cells(PFX - 2, PFY + PFW + 21)).Interior.Color = GamSheBC '(1,21)~(1,32)
    Range(Cells(PFX + PFH + 1, PFY + 18), Cells(PFX + PFH + 1, PFY + PFW + 21)).Interior.Color = GamSheBC ' (20,21)~(20,32)
    Range(Cells(PFX - 2, PFY + 18), Cells(PFX + PFH + 1, PFY + PFW + 21)).Interior.Color = GamSheBC '(1,21)~(20,21)
    Range(Cells(PFX - 2, PFY + 18), Cells(PFX + PFH + 1, PFY + PFW + 21)).Interior.Color = GamSheBC '(1,32)~(20,32)
   
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
    
    '2p
    With Range(Cells(PFX - 1, PFY + 19), Cells(PFX + PFH, PFY + PFW + 20))
        .Borders(8).Color = PFBBC  ' Top border
        .Borders(8).Weight = 4
        .Borders(7).Color = PFBBC  ' Bottom border
        .Borders(7).Weight = 4
        .Borders(9).Color = PFBDC  ' Left border
        .Borders(9).Weight = 4
        .Borders(10).Color = PFBDC  ' Right border
        .Borders(10).Weight = 4
    End With
    '2p
    Range(Cells(PFX - 1, PFY + 19), Cells(PFX + PFH, PFY + PFW + 20)).Interior.Color = PFBNC
    
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
    
    ' 2p
    With Range(Cells(PFX, PFY + 20), Cells(PFX + PFH - 1, PFY + PFW + 19))
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
    
    
End Sub


Sub UpdateGame_2p()
    ' Initialize game by setting initial values, preparing the game, creating the sheet
    Call SetInitialValues
    Call InitializeGame
    Call CreateGameSheet_2p
End Sub
