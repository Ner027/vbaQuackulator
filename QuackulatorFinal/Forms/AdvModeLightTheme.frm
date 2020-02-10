VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AdvModeLightTheme 
   Caption         =   "Quackulator"
   ClientHeight    =   7956
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5808
   OleObjectBlob   =   "AdvModeLightTheme.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AdvModeLightTheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PublicVars
Dim num(1 To 2) As Double
Dim optype As String
Dim res As Double
Dim opcount As Integer
Dim eqPressed As Boolean
Dim fontSize As Integer

Private Sub bte_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ansBox.Visible = True
End Sub

Private Sub CommandButton1_Click()
    Call advCalc("%")
End Sub

Private Sub CommandButton1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call advCalc("sqr")
End Sub

Private Sub CommandButton5_Click()
    Dim rLen As Integer
    If Len(resBox2.Text) = 1 And InStr(resBox2.Text, "0") And InStr(resBox2.Text, "-") = 0 Then
        Call key("-")
    End If
    If InStr(resBox2.Text, "-") = 0 Then
        Call key("-")
    End If
End Sub

'startup
Sub UserForm_Initialize()
    'Makes the Answers combobox invisible on startup
    ansBox.Visible = False
    'calls the main sub in the module to find the path to Sfx's
    Call Main
    th = 2
    md = 2
    'If options menu skiped
    'Sets number of floats to 7
    If dec = 0 Then
        dec = 7
    End If
    'CountersStartUp
    eqPressed = False
    resBox1.Text = ""
    resBox2.Text = mRes
    opcount = 0
End Sub

'KeypadFunction
Function key(a As String)
    Call playSound
    If Len(resBox2.Text) <= 13 Then
        If resBox2.Text = "0" Then
            resBox2.Text = a
        Else
            resBox2.Text = resBox2.Text + a
        End If
    End If
End Function

'Buttons
Private Sub btp_Click()
    If InStr(1, resBox2.Text, ".") = 0 Then
        Call key(".")
    End If
End Sub
Private Sub bt0_Click()
    Call key(0)
End Sub

Private Sub bt1_Click()
    Call key(1)
End Sub

Private Sub bt2_Click()
    Call key(2)
End Sub

Private Sub bt3_Click()
    Call key(3)
End Sub

Private Sub bt4_Click()
    Call key(4)
End Sub

Private Sub bt5_Click()
    Call key(5)
End Sub

Private Sub bt6_Click()
    Call key(6)
End Sub

Private Sub bt7_Click()
    Call key(7)
End Sub

Private Sub bt8_Click()
    Call key(8)
End Sub

Private Sub bt9_Click()
    Call key(9)
End Sub

'OpButtons
Private Sub btdiv_Click()
    Call op("/")
End Sub
Private Sub btmin_Click()
        Call op("-")
End Sub
Private Sub btplus_Click()
    Call op("+")
End Sub

Private Sub btx_Click()
        Call op("x")
End Sub
Private Sub cosBt_Click()
    Call advCalc("cos")
End Sub

Private Sub sinBt_Click()
    Call advCalc("sin")
End Sub

Private Sub tanBt_Click()
    Call advCalc("tan")
End Sub
Private Sub bte_Click()
    If resBox1.Text <> "" Then
        Call playSound
        eqPressed = True
        Call eq
    End If
End Sub

'ClearAll
Private Sub cleard_Click()
    Call playSound
    Call UserForm_Initialize
    resBox2.Text = "0"
    resBox2.Font.Size = 33
End Sub

'BackSpace
Private Sub bksp_Click()
    Call playSound
    Dim strLen As Integer
    strLen = Len(resBox2.Text)
    
    If strLen > 0 And Left(resBox2.Text, 1) <> "0" Then
        resBox2.Text = Left(resBox2.Text, (strLen - 1))
    End If
End Sub

'ClearEntry
Private Sub clearent_Click()
    Call playSound
    resBox2.Text = "0"
    resBox2.Font.Size = 33
End Sub

'GotoOptionsMenu
Private Sub cleard_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    mRes = res
    Options.ComboBox1.Value = dec
    If quack = True Then
        Options.CheckBox1.Value = True
    End If
    AdvModeLightTheme.Hide
    Options.Show
End Sub

'Offbutton
Private Sub off_Click()
    Unload Me
End Sub

'OperationButtonsRoutine
Sub op(b As String)
    If resBox2.Font.Size < 36 Then
        resBox1.Font.Size = resBox2.Font.Size - 7
    End If
        resBox2.Font.Size = 33
    Call playSound
    opcount = opcount + 1
    'For first operation being performed
    If opcount <= 1 Then
        num(1) = Val(resBox2.Text)
        resBox1.Text = Str(num(1)) & b
        resBox2.Text = ""
    'If it's not first operation and equal button was not pressed
    ElseIf opcount > 1 And eqPressed = False Then
        'If a second number was not typed,value will be the same only operation
        'is going to be changed
        If resBox2.Text = "" Then
            num(1) = Val(resBox1.Text)
            resBox1.Text = Str(num(1)) & b
        Else
        'If a second number was typed,will perform the operation
        'and print the result
            Call eq
            num(1) = Val(resBox2.Text)
            resBox1.Text = Str(num(1)) & b
            resBox2.Text = ""
        End If
    'If it's not the first operation
    'but equal button has been pressed
    ElseIf opcount > 1 And eqPressed = True Then
        num(1) = Val(resBox2.Text)
        resBox1.Text = Str(num(1)) & b
        resBox2.Text = ""
    End If
    eqPressed = False
    optype = b
End Sub

'EqualRoutine
Sub eq()
    num(2) = Val(resBox2.Text)
    'Calculates result
    'Round is used to round the numbers to
    'the float number defined in the options
    'if none was defined then it is = 7 by default
    Select Case optype
        Case "+"
            res = Round((num(1) + num(2)), dec)
        Case "-"
            res = Round((num(1) - num(2)), dec)
        Case "x"
            res = Round((num(1) * num(2)), dec)
        Case "/"
            If num(2) <> 0 Then
                res = Round((num(1) / num(2)), dec)
            Else
            'If someone tries to divide by zero
            'shows a warning to prevent the world from exploding
                MsgBox "Can't Divide By Zero"
            End If
    End Select
    'In case the number is to big to prevents the display
    'from exploding the following code changes the font size
    Call fontSizeCalc
    If fontSize < 36 Then
        resBox2.Font.Size = fontSize
    End If
    'Updates the displays
    resBox2.Text = Str(res)
    resBox1.Text = ""
    'Resets the numbers
    num(1) = 0
    num(2) = 0
    Call updateAns
End Sub
'advCalc
Sub advCalc(t As String)
    Select Case t
        Case "sin"
            If resBox2.Text <> "" And resBox1.Text = "" Then
                resBox2.Text = Str(Round(Math.Sin((Val(resBox2.Text) * 3.14159265359) / 180), dec))
            ElseIf resBox2.Text <> "" And resBox1.Text <> "" Then
                Call eq
                resBox2.Text = Str(Round(Math.Sin((Val(resBox2.Text) * 3.14159265359) / 180), dec))
            End If
        Case "cos"
            If resBox2.Text <> "" And resBox1.Text = "" Then
                resBox2.Text = Str(Round(Math.Cos((Val(resBox2.Text) * 3.14159265359) / 180), dec))
            ElseIf resBox2.Text <> "" And resBox1.Text <> "" Then
                Call eq
                resBox2.Text = Str(Round(Math.Cos((Val(resBox2.Text) * 3.14159265359) / 180), dec))
            End If
        Case "tan"
            If resBox2.Text <> "" And resBox1.Text = "" Then
                resBox2.Text = Str(Round(Math.Tan((Val(resBox2.Text) * 3.14159265359) / 180), dec))
            ElseIf resBox2.Text <> "" And resBox1.Text <> "" Then
                Call eq
                resBox2.Text = Str(Round(Math.Tan((Val(resBox2.Text) * 3.14159265359) / 180), dec))
            End If
        Case "%"
            If resBox2.Text <> "" And resBox1.Text = "" Then
                resBox2.Text = Str(Round(Val(resBox2.Text) * 0.01, dec))
            ElseIf resBox2.Text <> "" And resBox1.Text <> "" Then
                Call eq
                resBox2.Text = Str(Round(Val(resBox2.Text) * 0.01, dec))
            End If
        Case "sqr"
            If resBox2.Text <> "" And resBox1.Text = "" Then
                resBox2.Text = Str(Round(Val(resBox2.Text) * 0.01, dec))
            ElseIf resBox2.Text <> "" And resBox1.Text <> "" Then
                Call eq
                resBox2.Text = Str(Round(Sqr(Val(resBox2.Text)), dec))
            End If
    End Select
End Sub
'CalculatesMaxFontSize
Sub fontSizeCalc()
    Dim maxW As Long
    maxW = (228 / Len(Str(res)))
    fontSize = Int((maxW * 1.7) + 3)
    Debug.Print fontSize
End Sub

'Answers

'UpdatesTheHistory
Sub updateAns()
    ansBox.AddItem Str(res)
End Sub

'Locks the results ComboBox so that the user
'Can't to any stupid thing
Private Sub ansBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub

'When the user picks a previous value it changes
'the input to that value
Private Sub ansBox_Change()
    resBox2.Text = ansBox.Value
    ansBox.Visible = False
End Sub

