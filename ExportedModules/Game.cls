VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Worksheets("Game Records").Activate
End Sub

Private Sub CommandButton2_Click()
    UpdateGame
End Sub

Private Sub CommandButton3_Click()
    StartNewGame
End Sub

Private Sub CommandButton4_Click()
    PauseAndRestart
End Sub

Private Sub CommandButton5_Click()
    OpenGithub
End Sub

Private Sub CommandButton6_Click()
    RegenerateNextBlockAndContinue
End Sub

Private Sub CommandButton7_Click()
    QuitTheGame
End Sub
