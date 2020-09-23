Attribute VB_Name = "modPlaySound"
Option Explicit

'Private Declare Function PlaySoundData Lib "winmm.dll" Alias "PlaySoundA" _
'    (lpData As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long
'
'Private m_snd() As Byte
'
'Public Enum WAVInRes
'    wirBanker = 101
'    wirTension = 102
'    wirTheme = 103
'    wirStop = 0
'    wir
'End Enum
'
'Public Enum PlayType
'    SND_SYNC = &H0 ' play synchronously (default)
'    SND_ASYNC = &H1 ' play asynchronously
'    SND_NODEFAULT = &H2 ' silence not default, if sound not found
'    SND_MEMORY = &H4 ' lpszSoundName points to a memory file
'    SND_ALIAS = &H10000 ' name is a WIN.INI [sounds] entry
'    SND_FILENAME = &H20000 ' name is a file name
'    SND_RESOURCE = &H40004 ' name is a resource name or atom
'    SND_ALIAS_ID = &H110000 ' name is a WIN.INI [sounds] entry identifier
'    SND_ALIAS_START = 0 ' must be > 4096 to keep strings in same section of resource file
'    SND_LOOP = &H8 ' loop the sound until next sndPlaySound
'    SND_NOSTOP = &H10 ' don't stop any currently playing sound
'    SND_VALID = &H1F ' valid flags /;Internal /
'    SND_NOWAIT = &H2000 ' don't wait if the driver is busy
'    SND_VALIDFLAGS = &H17201F ' Set of valid flag bits. Anything outside this range will raise error
'    SND_RESERVED = &HFF000000 ' In particular these flags are reserved
'    SND_TYPE_MASK = &H170007
'End Enum

Public PName As String, Vol As Integer, Offered As New Collection, Mute As Boolean
Public HSName() As String, HSScore() As String, CurScore As Long

'Public Function PlayResSoundData(ByVal SndID As Integer, Optional ByVal prsPlayType As PlayType) As Long
'Const Flags = prsPlayType 'SND_MEMORY Or SND_ASYNC Or SND_NODEFAULT
'If SndID = 0 Then PlaySoundData "", 0, SND_NODEFAULT Or SND_ASYNC Or SND_MEMORY Or SND_ALIAS_START: Exit Function
'm_snd = LoadResData(SndID, "CUSTOM")
'PlaySoundData m_snd(0), 0, prsPlayType
'End Function

Public Sub PlayMusic(ByVal Index As Byte)
frmSplash.spSound.StopSound 1
frmSplash.spSound.PlaySound App.Path & "\Music\dondM" & Index & ".mp3", True
End Sub

Public Sub StopMusic()
frmSplash.spSound.StopAll
End Sub

Public Sub PlaySE(ByVal Index As Byte)
frmSplash.spSound.PlaySound App.Path & "\Sound FX\dondSE" & Index & ".mp3", False
End Sub

Private Function tName() As String
tName = Join(HSName, "ÿ")
End Function

Private Function tScore() As String
tScore = Join(HSScore, "ÿ")
End Function

Public Sub SaveHighScores()
Dim hPath As String, i As Integer
hPath = App.Path & "\DealOrNoDeal.hs"
For i = 0 To 2
    If Val(CurScore) > Val(HSScore(i)) Then
        HSScore(i) = CurScore
        HSName(i) = PName
        Exit For
    End If
    'MsgBox UBound(HSScore)
Next i
Open hPath For Random As #1
    Put #1, 1, tName
    Put #1, 2, tScore
Close #1
End Sub

Public Sub LoadHighScores()
Dim hPath As String, i As Integer
hPath = App.Path & "\DealOrNoDeal.hs"
If Dir(hPath) = "" Then
    Open hPath For Random As #1
        ReDim HSName(2)
        ReDim HSScore(2)
        For i = 0 To 2
            HSName(i) = "Player 1"
            HSScore(i) = "0"
        Next i
        Put #1, 1, tName
        Put #1, 2, tScore
    Close #1
End If
Open hPath For Random As #1
    Dim sTempN As String, sTempS As String
    Get #1, 1, sTempN
    Get #1, 2, sTempS
    HSName = Split(sTempN, "ÿ")
    HSScore = Split(sTempS, "ÿ")
Close #1
End Sub
