VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModeSelection 
   Caption         =   "UserForm1"
   ClientHeight    =   8232.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8580.001
   OleObjectBlob   =   "ModeSelection.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "ModeSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    
    Game.Visible = xlSheetVisible
    Game2P.Visible = xlSheetVisible
    Worksheets("Game").Activate
    Game2P.Visible = xlSheetVeryHidden
    Call SetInitialValues
    Call InitializeGame
    Call CreateGameSheet
    Unload ModeSelection
    
End Sub

Private Sub CommandButton2_Click()
    
    Game.Visible = xlSheetVisible
    Game2P.Visible = xlSheetVisible
    Worksheets("2p").Activate
    Game.Visible = xlSheetVeryHidden
    Call SetInitialValues
    Call InitializeGame_2p
    Call CreateGameSheet_2p
    Unload ModeSelection
    
End Sub

Private Sub CommandButton3_Click()
    For Each W In Application.Workbooks
    W.Save
    Next W
    Application.Quit
End Sub

Private Sub CommandButton4_Click()
    Rule.Show
End Sub
