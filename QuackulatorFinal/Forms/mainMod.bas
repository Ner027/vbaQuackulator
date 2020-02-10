Attribute VB_Name = "mainMod"
Option Explicit
'MultiFormVars
Public th As Integer
Public md As Integer
Public dec As Integer
Public quack As Boolean
Public mRes As String
Public smenuPath As String
Public nullSound As String
'ModuleVars
Dim path As String
Dim fullPath As String
Public Declare PtrSafe Function sndPlaySound Lib "winmm.dll" _
Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
ByVal uFlags As Long) As Long
Sub Main()
    path = ActiveDocument.path
    fullPath = path + "\Resources\Soundbank\sfx\quack.wav"
    smenuPath = path + "\Resources\Soundbank\sfx\startMenu.wav"
    nullSound = path + "\Resources\Soundbank\sfx\null.wav"
    Debug.Print fullPath
End Sub

Sub playSound()
    If quack = True Then
        sndPlaySound fullPath, 1
    End If
End Sub
