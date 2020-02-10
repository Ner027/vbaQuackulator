VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} easterEggs 
   Caption         =   "Quackulator"
   ClientHeight    =   6072
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10788
   OleObjectBlob   =   "easterEggs.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "easterEggs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    easterEggs.Hide
    Piano.Show
End Sub

Private Sub CommandButton2_Click()
    easterEggs.Hide
    drumpad.Show
End Sub


