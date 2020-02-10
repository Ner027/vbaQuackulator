VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Piano 
   Caption         =   "Quackulator"
   ClientHeight    =   4164
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7416
   OleObjectBlob   =   "Piano.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Piano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Function sndPlaySound Lib "winmm.dll" _
Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
ByVal uFlags As Long) As Long

Dim bankPath As String
Dim notePath As String
Dim octave As Integer

Private Sub CommandButton1_Click()
    Call note("C")
End Sub

Private Sub CommandButton10_Click()
    Call note("B")
End Sub

Private Sub CommandButton11_Click()
    Call note("G#")
End Sub

Private Sub CommandButton12_Click()
    Call note("A#")
End Sub

Private Sub CommandButton13_Click()
    mciExecute "play" & "C:\Users\necos\OneDrive\Ambiente de Trabalho\Quackulator\fur.mid"
End Sub

Private Sub CommandButton2_Click()
    Call note("D")
End Sub

Private Sub CommandButton3_Click()
    Call note("E")
End Sub

Private Sub CommandButton4_Click()
    Call note("F")
End Sub

Private Sub CommandButton5_Click()
    Call note("C#")
End Sub

Private Sub CommandButton6_Click()
    Call note("D#")
End Sub

Private Sub CommandButton7_Click()
    Call note("G")
End Sub

Private Sub CommandButton8_Click()
    Call note("F#")
End Sub

Private Sub CommandButton9_Click()
    Call note("A")
End Sub

Private Sub oct_Change()
    octave = oct.Value
End Sub
Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim key As String
    Debug.Print Chr(KeyAscii)
    key = LCase(Chr(KeyAscii))
    TextBox1.Text = ""
    Select Case key
        Case "q"
            Call note("C")
        Case "2"
            Call note("C#")
        Case "w"
            Call note("D")
         Case "3"
            Call note("D#")
        Case "e"
            Call note("E")
        Case "r"
            Call note("F")
        Case "5"
            Call note("F#")
        Case "t"
            Call note("G")
        Case "6"
            Call note("G#")
        Case "y"
            Call note("A")
        Case "7"
            Call note("A#")
        Case "u"
            Call note("B")
    End Select
End Sub

Private Sub UserForm_Initialize()
    octave = 5
    oct.Value = octave
    oct.AddItem 4
    oct.AddItem 5
    oct.AddItem 6
    bankPath = ActiveDocument.path + "\Resources\Soundbank\Piano\"
End Sub

Sub note(n As String)
    notePath = bankPath & n & octave & ".wav"
    sndPlaySound notePath, 1
    Debug.Print notePath
End Sub


