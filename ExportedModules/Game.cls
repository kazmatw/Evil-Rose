VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
<<<<<<< Updated upstream
=======
Private Sub CommandButton1_Click()
    
    GameRecords.Visible = xlSheetVisible
    Worksheets("Game Records").Activate
    Game.Visible = xlSheetVeryHidden
    
End Sub

Private Sub CommandButton10_Click()
    BGM_Next
End Sub

Private Sub CommandButton11_Click()
    BGM_Prev
End Sub

Private Sub CommandButton2_Click()
    UpdateGame
End Sub

Private Sub CommandButton3_Click()
    StartNewGame
End Sub

Private Sub CommandButton4_Click()
    Rule.Show
End Sub

Private Sub CommandButton5_Click()
    OpenGithub
End Sub


Private Sub CommandButton6_Click()
    ImportModules
End Sub

Private Sub CommandButton7_Click()
    QuitTheGame
End Sub

Private Sub CommandButton8_Click()
        
    Game2P.Visible = xlSheetVisible
    Worksheets("2p").Activate
    Game.Visible = xlSheetVeryHidden
    
End Sub

Private Sub CommandButton9_Click()
    ExportModules
End Sub

Private Sub ScrollBar1_Change()
    SetVolume ScrollBar1.Value
End Sub
>>>>>>> Stashed changes
