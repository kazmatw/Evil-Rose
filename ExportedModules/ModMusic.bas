Attribute VB_Name = "ModMusic"
Private Declare PtrSafe Function PlaySound Lib "winmm.dll" _
        Alias "PlaySoundA" (ByVal lpszName As String, _
        ByVal hModule As LongPtr, ByVal dwFlags As Long) As Boolean


Public Sub BGM_Music()

    myPath = ActiveWorkbook.Path
    
    mA = myPath & "/Music/BGM.wav"
    
     Call PlaySound(mA, 0, 9)
     '(file path,beep sound 0 is false,1= Background Play 8=loop play 9=1+8)

End Sub


Public Sub pauseBGM()
    
     Call PlaySound(" ", 0, 1)
     'If there is no path will stop the music.

End Sub
