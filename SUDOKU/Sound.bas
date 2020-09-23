Attribute VB_Name = "Sound"
'Call API
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Const SND_ASYNC = &H1

'The sounds
Public Enum wHaTSoUnD
   Starting
   Done
End Enum

'The sub that plays the sounds
Public Sub PlaySnd(Soundtype As wHaTSoUnD)
    Select Case Soundtype
            Case Starting
            Call PlaySound(App.Path + "\Gong.wav", 0, SND_ASYNC)
            Case Done
            Call PlaySound(App.Path + "\Gong.wav", 0, SND_ASYNC)
                      
    End Select
    
End Sub

