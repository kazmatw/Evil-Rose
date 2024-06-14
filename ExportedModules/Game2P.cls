VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Game2P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub CommandButton1_Click()

    Call StartNewGame_2p

End Sub

Private Sub CommandButton2_Click()

    Call UpdateGame_2p

End Sub

Private Sub CommandButton3_Click()
    
    
    Game.Visible = xlSheetVisible
    Worksheets("Game").Activate
    Game2P.Visible = xlSheetVeryHidden
    GameRecords.Visible = xlSheetVeryHidden
End Sub
