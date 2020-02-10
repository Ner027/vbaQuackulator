VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Options 
   Caption         =   "Quackulator"
   ClientHeight    =   5436
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6780
   OleObjectBlob   =   "Options.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()
    If CheckBox1.Value = True Then
        quack = True
    ElseIf CheckBox1.Value = False Then
        quack = False
    End If
    Debug.Print quack
End Sub
Sub UserForm_Initialize()
    Select Case md
        Case Is = 1
        b.Value = True
        Case Is = 2
        a.Value = True
    End Select
    Select Case th
        Case Is = 1
        d.Value = True
        Case Is = 2
        l.Value = True
    End Select
Dim i As Integer
    For i = 1 To 10
        ComboBox1.AddItem i
    Next
End Sub
Private Sub CommandButton1_Click()
    If ComboBox1.Value <> "" Then
        dec = ComboBox1.Value
    Else
        dec = 7
    End If
    If th = 1 And md = 1 Then
        Options.Hide
        DarkMode.Show
    ElseIf th = 1 And md = 2 Then
        Options.Hide
        AdvModeDarkTheme.Show
    ElseIf th = 2 And md = 1 Then
        Options.Hide
        LightMode.Show
    ElseIf th = 2 And md = 2 Then
        Options.Hide
        AdvModeLightTheme.Show
    End If
End Sub

Private Sub d_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    th = 1
    Debug.Print th
    d.Value = True
    l.Value = False
End Sub
Private Sub l_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    th = 2
    l.Value = True
    d.Value = False
End Sub
Private Sub b_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    md = 1
    Debug.Print md
    b.Value = True
    a.Value = False
End Sub
Private Sub a_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    md = 2
    a.Value = True
    b.Value = False
End Sub
