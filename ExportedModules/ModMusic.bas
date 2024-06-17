Attribute VB_Name = "ModMusic"

Private Declare PtrSafe Function PlaySound Lib "winmm.dll" _
        Alias "PlaySoundA" (ByVal lpszName As String, _
        ByVal hModule As LongPtr, ByVal dwFlags As Long) As Boolean
   

        
Private Declare PtrSafe Function waveOutSetVolume Lib "winmm.dll" (ByVal hwo As LongPtr, ByVal dwVolume As Long) As Long


Public BGMList As Integer
Public AudioVolume As Long
Public PlayStop As Boolean

Sub BGM_Music()
 
    If BGMList = 0 Then BGMList = 1

    myPath = ActiveWorkbook.Path
    mA = myPath & "/Music/BGM" & CStr(BGMList) & ".wav"
    
     Call PlaySound(mA, 0, 9)
     '(file path,beep sound 0 is false,1= Background Play 8=loop play 9=1+8)
     
     PlayStop = True

End Sub

Sub pauseBGM()
    
     Call PlaySound(" ", 0, 1)
     'If there is no path will stop the music.
     PlayStop = False

End Sub

Sub SetVolume(volume As Long)
    
    'Make sure the volume is within range
    If volume < 0 Then volume = 0
    If volume > 65535 Then volume = 65535

    ' 0 is using the default audio device volume
    waveOutSetVolume 0, volume
    
End Sub

Sub VolumeUp()

    'Make sure the volume is within range
    AudioVolume = AudioVolume + (65535 * 0.1)
    If AudioVolume > 65535 Then AudioVolume = 65535
    
    ' 0 is using the default audio device volume
    waveOutSetVolume 0, AudioVolume

End Sub

Sub VolumeDown()
    
    'Make sure the volume is within range
    AudioVolume = AudioVolume - (65535 * 0.1)
    If AudioVolume < 0 Then AudioVolume = 0
    
    ' 0 is using the default audio device volume
    waveOutSetVolume 0, AudioVolume

End Sub

Sub BGM_Next()
    
   BGMList = BGMList + 1
   
   If BGMList > 10 Then BGMList = 1
   
   Call BGM_Music

End Sub

Sub BGM_Prev()
    
   BGMList = BGMList - 1
   
   If BGMList <= 0 Then BGMList = 10
   
   Call BGM_Music

End Sub


Sub BGM_Start()
    
   AudioVolume = 65535

   BGMList = 1
   
   Call BGM_Music
   
End Sub

Sub BGM_PlayorStop()
    
    If PlayStop = True Then
        Call pauseBGM
    Else
        Call BGM_Music
    End If

End Sub


Sub GameOverSoundEffect()

    myPath = ActiveWorkbook.Path
    mA = myPath & "/Music/gameover.wav"
    
    Call PlaySound(mA, 0, 1)
     
    PlayStop = False
    
End Sub
