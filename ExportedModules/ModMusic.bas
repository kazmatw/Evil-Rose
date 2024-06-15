Attribute VB_Name = "ModMusic"

Private Declare PtrSafe Function PlaySound Lib "winmm.dll" _
        Alias "PlaySoundA" (ByVal lpszName As String, _
        ByVal hModule As LongPtr, ByVal dwFlags As Long) As Boolean
   

        
Private Declare PtrSafe Function waveOutSetVolume Lib "winmm.dll" (ByVal hwo As LongPtr, ByVal dwVolume As Long) As Long


Public BGMList As Integer

Public Sub BGM_Music()

    myPath = ActiveWorkbook.Path
    mA = myPath & "/Music/BGM" & CStr(BGMList) & ".wav"
    
     Call PlaySound(mA, 0, 9)
     '(file path,beep sound 0 is false,1= Background Play 8=loop play 9=1+8)

End Sub


Public Sub pauseBGM()
    
     Call PlaySound(" ", 0, 1)
     'If there is no path will stop the music.

End Sub
<<<<<<< Updated upstream
=======

Sub SetVolume(volume As Long)
    
    'Make sure the volume is within range
    If volume < 0 Then volume = 0
    If volume > 65535 Then volume = 65535

    ' 0 is using the default audio device volume
    waveOutSetVolume 0, volume
    
    
End Sub

Sub BGM_Next()
    
   BGMList = BGMList + 1
   
   
   If BGMList > 3 Then BGMList = 1
   
   Call BGM_Music

End Sub

Sub BGM_Prev()
    
   BGMList = BGMList - 1
   
   If BGMList <= 0 Then BGMList = 3
   
   Call BGM_Music

End Sub


Sub BGM_Start()
    
   BGMList = 1
   
   Call BGM_Music

End Sub





>>>>>>> Stashed changes
