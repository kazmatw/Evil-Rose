Attribute VB_Name = "Modscene"
Sub CoverToMenu()
    Call ClickSoundEffect
    Menu.Visible = xlSheetVisible
    Sheets("Menu").Activate
    Cover.Visible = xlSheetVeryHidden
End Sub

Sub MenuToCover()
    Call ClickSoundEffect
    Cover.Visible = xlSheetVisible
    Sheets("Cover").Activate
    Menu.Visible = xlSheetVeryHidden
End Sub

Sub MenuTo1p()
    Call ClickSoundEffect
    Game.Visible = xlSheetVisible
    Sheets("Game").Activate
    Menu.Visible = xlSheetVeryHidden
End Sub

Sub MenuTo2p()
    Call ClickSoundEffect
    Game2P.Visible = xlSheetVisible
    Sheets("Game2p").Activate
    Menu.Visible = xlSheetVeryHidden
End Sub

Sub MenuToRules()
    Call ClickSoundEffect
    Rules.Visible = xlSheetVisible
    Sheets("Rules").Activate
    Menu.Visible = xlSheetVeryHidden
End Sub

Sub MenuToLeaderboard()
    Call ClickSoundEffect
    Record.Visible = xlSheetVisible
    Sheets("Record").Activate
    Menu.Visible = xlSheetVeryHidden
End Sub

Sub MenuToComingsoon()
    Call ClickSoundEffect
    ComingSoon.Visible = xlSheetVisible
    Sheets("Comingsoon").Activate
    Menu.Visible = xlSheetVeryHidden
End Sub

Sub MenuToMusic()
    Call ClickSoundEffect
    Music.Visible = xlSheetVisible
    Sheets("Music").Activate
    Menu.Visible = xlSheetVeryHidden
End Sub

Sub MusicToMenu()
    Call ClickSoundEffect
    Menu.Visible = xlSheetVisible
    Sheets("Menu").Activate
    Music.Visible = xlSheetVeryHidden
End Sub

Sub ComingsoonToMenu()
    Call ClickSoundEffect
    Menu.Visible = xlSheetVisible
    Sheets("Menu").Activate
    ComingSoon.Visible = xlSheetVeryHidden
End Sub

Sub LeaderboardToMenu()
    Call ClickSoundEffect
    Menu.Visible = xlSheetVisible
    Sheets("Menu").Activate
    Record.Visible = xlSheetVeryHidden
End Sub

Sub RulesToMenu()
    Call ClickSoundEffect
    Menu.Visible = xlSheetVisible
    Sheets("Menu").Activate
    Rules.Visible = xlSheetVeryHidden
End Sub

Sub GameToMenu()
    Call ClickSoundEffect
    Menu.Visible = xlSheetVisible
    Sheets("Menu").Activate
    Game.Visible = xlSheetVeryHidden
End Sub

Sub Game2pToMenu()
    Call ClickSoundEffect
    Menu.Visible = xlSheetVisible
    Sheets("Menu").Activate
    Game2P.Visible = xlSheetVeryHidden
End Sub

