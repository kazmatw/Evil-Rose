VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Leaderboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Game.Visible = xlSheetVisible
    Worksheets("Game").Activate
    GameRecords.Visible = xlSheetVeryHidden
End Sub

Private Sub CommandButton2_Click()
    Call ClearHistory
End Sub
