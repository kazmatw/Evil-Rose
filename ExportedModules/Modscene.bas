Attribute VB_Name = "Modscene"
Sub CoverToMenu()
    Menu.Visible = xlSheetVisible
    Sheets("Menu").Activate
    Cover.Visible = xlSheetVeryHidden
End Sub

Sub MenuToCover()
    Cover.Visible = xlSheetVisible
    Sheets("Cover").Activate
    Menu.Visible = xlSheetVeryHidden
End Sub

Sub MenuTo1p()
    Game.Visible = xlSheetVisible
    Sheets("Game").Activate
    Menu.Visible = xlSheetVeryHidden
End Sub

Sub MenuTo2p()
    Game2P.Visible = xlSheetVisible
    Sheets("Game2p").Activate
    Menu.Visible = xlSheetVeryHidden
End Sub

Sub MenuToRules()
    Rules.Visible = xlSheetVisible
    Sheets("Rules").Activate
    Menu.Visible = xlSheetVeryHidden
End Sub

Sub MenuToLeaderboard()
    Record.Visible = xlSheetVisible
    Sheets("Record").Activate
    Menu.Visible = xlSheetVeryHidden
End Sub

Sub MenuToComingsoon()
    ComingSoon.Visible = xlSheetVisible
    Sheets("Comingsoon").Activate
    Menu.Visible = xlSheetVeryHidden
End Sub

Sub ComingsoonToMenu()
    Menu.Visible = xlSheetVisible
    Sheets("Menu").Activate
    ComingSoon.Visible = xlSheetVeryHidden
End Sub

Sub LeaderboardToMenu()
    Menu.Visible = xlSheetVisible
    Sheets("Menu").Activate
    Record.Visible = xlSheetVeryHidden
End Sub

Sub RulesToMenu()
    Menu.Visible = xlSheetVisible
    Sheets("Menu").Activate
    Rules.Visible = xlSheetVeryHidden
End Sub

Sub GameToMenu()
    Menu.Visible = xlSheetVisible
    Sheets("Menu").Activate
    Game.Visible = xlSheetVeryHidden
End Sub

Sub Game2pToMenu()
    Menu.Visible = xlSheetVisible
    Sheets("Menu").Activate
    Game2P.Visible = xlSheetVeryHidden
End Sub

