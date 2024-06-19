Attribute VB_Name = "ModMusic"
Private Declare PtrSafe Function PlaySound Lib "winmm.dll" _
        Alias "PlaySoundA" (ByVal lpszName As String, _
        ByVal hModule As LongPtr, ByVal dwFlags As Long) As Boolean

Private Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, _
     ByVal lpstrReturnString As String, _
     ByVal uReturnLength As Long, _
     ByVal hwndCallback As Long) As Long

Public BGMList As Integer
Public AudioVolume As Long
Public command As String
Public PlayStop As Boolean

Public Sub BGM_Music()

    If BGMList = 0 Then BGMList = 1
    
    myPath = ActiveWorkbook.Path
    command = "open """ & myPath & "\Music\BGM" & CStr(BGMList) & ".mp3"" type mpegvideo alias MyMusic"
    Call mciSendString(command, vbNullString, 0, 0)
    'open the music file
    
    command = "play MyMusic repeat"
    ' repeat the BGM
    Call mciSendString(command, vbNullString, 0, 0)
    
    PlayStop = True
     
End Sub

Sub pauseBGM()
    
    mciSendString "stop MyMusic", vbNullString, 0, 0
    mciSendString "close MyMusic", vbNullString, 0, 0
    PlayStop = False

End Sub

Sub realpauseBGM()
    
    mciSendString "pause MyMusic", vbNullString, 0, 0
    PlayStop = False

End Sub

Sub VolumeUp()

     'Make sure the volume is within range
    AudioVolume = AudioVolume + (1000 * 0.1)
    If AudioVolume > 1000 Then AudioVolume = 1000
    
    mciSendString "setaudio MyMusic volume to " & AudioVolume, vbNullString, 0, 0

End Sub

Sub VolumeDown()
    
    'Make sure the volume is within range
    AudioVolume = AudioVolume - (1000 * 0.1)
    If AudioVolume < 0 Then AudioVolume = 0
    
    mciSendString "setaudio MyMusic volume to " & AudioVolume, vbNullString, 0, 0

End Sub

Sub BGM_Next()
    
   BGMList = BGMList + 1
   
   If BGMList > 5 Then BGMList = 1
   
   Call pauseBGM
   Call BGM_Music

End Sub

Sub BGM_Prev()
    
   BGMList = BGMList - 1
   
   If BGMList <= 0 Then BGMList = 5
   
   Call pauseBGM
   Call BGM_Music

End Sub


Sub BGM_Start()

   BGMList = 1
   
   Call BGM_Music
   
   AudioVolume = 100
   mciSendString "setaudio MyMusic volume to " & AudioVolume, vbNullString, 0, 0
   
End Sub

Sub BGM_PlayorStop()
    
    If PlayStop = True Then
        Call realpauseBGM
    Else
        Call BGM_Music
    End If

End Sub


Sub GameOverSoundEffect()

    myPath = ActiveWorkbook.Path
    mA = myPath & "/Music/gameover.wav"
    
    Call PlaySound(mA, 0, 1)
    
End Sub

Sub ClickSoundEffect()

    myPath = ActiveWorkbook.Path
    mA = myPath & "/Music/ClickEffect.wav"
    
    Call PlaySound(mA, 0, 1)
    
End Sub

Sub DeleteSoundEffect()

    myPath = ActiveWorkbook.Path
    mA = myPath & "/Music/deleteBlock.wav"
    
    Call PlaySound(mA, 0, 1)
    
End Sub

Sub ComboDeleteSoundEffect()

    myPath = ActiveWorkbook.Path
    mA = myPath & "/Music/ComboDelete.wav"
    
    Call PlaySound(mA, 0, 1)
    
End Sub

Sub DeleteRowSoundEffect()

    myPath = ActiveWorkbook.Path
    mA = myPath & "/Music/DeleteRow.wav"
    
    Call PlaySound(mA, 0, 1)
    
End Sub

