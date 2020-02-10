VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StartMenu 
   Caption         =   "Quackulator"
   ClientHeight    =   6072
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10788
   OleObjectBlob   =   "StartMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StartMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim browserPath As String
Dim portablePath As String
'StartButtonPress
Private Sub CommandButton1_Click()
    sndPlaySound nullSound, 1
    startMenu.Hide
    DarkMode.Show
End Sub

'OptionsButtonPress
Private Sub CommandButton2_Click()
    startMenu.Hide
    Options.Show
    sndPlaySound nullSound, 1
End Sub

'AboutButton
Private Sub CommandButton3_Click()
    sndPlaySound nullSound, 1
    'Needs a browser path to work
    If browserPath = VBA.Constants.vbNullString Then
        MsgBox "Broser not found!"
    Else
        Shell browserPath & " -URL " & "https://github.com/Ner027/vbaQuackulator"
    End If
End Sub

'MuteButton
Private Sub CommandButton4_Click()
    sndPlaySound nullSound, 1
End Sub


Sub UserForm_Initialize()
    browserPath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    portablePath = ActiveDocument.path + "\Resources\Chrome\GoogleChromePortable\chrome.exe"
    mRes = 0
    th = 0
    md = 0
    Call Main
    Call sndPlaySound(smenuPath, 1)
End Sub

Private Sub xD_Click()
    sndPlaySound nullSound, 1
    startMenu.Hide
    easterEggs.Show
End Sub
