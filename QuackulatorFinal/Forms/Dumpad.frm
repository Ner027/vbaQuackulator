VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dumpad 
   Caption         =   "Quackulator"
   ClientHeight    =   3972
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3468
   OleObjectBlob   =   "Dumpad.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Dumpad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Function sndPlaySound Lib "winmm.dll" _
Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
ByVal uFlags As Long) As Long
Dim samplePath As String

Private Sub CommandButton2_Click()
    Call play("snare")
End Sub

Private Sub CommandButton3_Click()
    Call play("ride")
End Sub

Private Sub CommandButton4_Click()
    Call play("tom")
End Sub

Private Sub CommandButton5_Click()
    Call play("hho")
End Sub

Private Sub CommandButton6_Click()
    Call play("hhc")
End Sub

Private Sub CommandButton7_Click()
    Call play("hit2")
End Sub

Private Sub CommandButton8_Click()
    Call play("crash")
End Sub

Private Sub CommandButton9_Click()
    Call play("hit")
End Sub
Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim drum As String
    TextBox1.Text = ""
    drum = LCase(Chr(KeyCode))
    Select Case drum
        Case "t"
            Call play("crash")
        Case "y"
            Call play("hit")
        Case "u"
            Call play("hit2")
        Case "g"
            Call play("hho")
        Case "h"
            Call play("hhc")
        Case "j"
            Call play("tom")
        Case "v"
            Call play("kick")
        Case "b"
            Call play("snare")
        Case "n"
            Call play("ride")
    End Select
End Sub

Sub UserForm_Initialize()
    samplePath = ActiveDocument.path + "\Resources\Soundbank\Drums\"
End Sub
Private Sub CommandButton1_Click()
    Call play("kick")
End Sub

Sub play(d As String)
    Call sndPlaySound(samplePath & d & ".wav", 1)
End Sub
