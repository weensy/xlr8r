VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormCP 
   Caption         =   "Color Picker"
   ClientHeight    =   1470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   1950
   OleObjectBlob   =   "FormCP.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "FormCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private turnedOn    As Object

Private Sub TurnOff(lblOff As Object)
    If Not turnedOn Is Nothing Then
        lblOff.BorderColor = RGB(100, 100, 100)
        
        'To avoid bug of excel
        lblOff.BorderStyle = fmBorderStyleNone
        lblOff.BorderStyle = fmBorderStyleSingle
    End If
End Sub

Private Sub TurnOn(lblOn As Object)
    lblOn.BorderColor = RGB(255, 255, 255)
    Set turnedOn = lblOn
End Sub

Private Sub lblB1_Click()
    objCaller.BackColor = lblB1.BackColor
    Unload Me
End Sub

Private Sub lblB1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblB1)
End Sub

Private Sub lblB2_Click()
    objCaller.BackColor = lblB2.BackColor
    Unload Me
End Sub

Private Sub lblB2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblB2)
End Sub

Private Sub lblB3_Click()
    objCaller.BackColor = lblB3.BackColor
    Unload Me
End Sub

Private Sub lblB3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblB3)
End Sub

Private Sub lblB4_Click()
    objCaller.BackColor = lblB4.BackColor
    Unload Me
End Sub

Private Sub lblB4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblB4)
End Sub

Private Sub lblB5_Click()
    objCaller.BackColor = lblB5.BackColor
    Unload Me
End Sub

Private Sub lblB5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblB5)
End Sub

Private Sub lblB6_Click()
    objCaller.BackColor = lblB6.BackColor
    Unload Me
End Sub

Private Sub lblB6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblB6)
End Sub

Private Sub lblG1_Click()
    objCaller.BackColor = lblG1.BackColor
    Unload Me
End Sub

Private Sub lblG1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblG1)
End Sub

Private Sub lblG2_Click()
    objCaller.BackColor = lblG2.BackColor
    Unload Me
End Sub

Private Sub lblG2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblG2)
End Sub

Private Sub lblG3_Click()
    objCaller.BackColor = lblG3.BackColor
    Unload Me
End Sub

Private Sub lblG3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblG3)
End Sub

Private Sub lblG4_Click()
    objCaller.BackColor = lblG4.BackColor
    Unload Me
End Sub

Private Sub lblG4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblG4)
End Sub

Private Sub lblG5_Click()
    objCaller.BackColor = lblG5.BackColor
    Unload Me
End Sub

Private Sub lblG5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblG5)
End Sub

Private Sub lblG6_Click()
    objCaller.BackColor = lblG6.BackColor
    Unload Me
End Sub

Private Sub lblG6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblG6)
End Sub

Private Sub lblO1_Click()
    objCaller.BackColor = lblO1.BackColor
    Unload Me
End Sub

Private Sub lblO1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblO1)
End Sub

Private Sub lblO2_Click()
    objCaller.BackColor = lblO2.BackColor
    Unload Me
End Sub

Private Sub lblO2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblO2)
End Sub

Private Sub lblO3_Click()
    objCaller.BackColor = lblO3.BackColor
    Unload Me
End Sub

Private Sub lblO3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblO3)
End Sub

Private Sub lblO4_Click()
    objCaller.BackColor = lblO4.BackColor
    Unload Me
End Sub

Private Sub lblO4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblO4)
End Sub

Private Sub lblO5_Click()
    objCaller.BackColor = lblO5.BackColor
    Unload Me
End Sub

Private Sub lblO5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblO5)
End Sub

Private Sub lblO6_Click()
    objCaller.BackColor = lblO6.BackColor
    Unload Me
End Sub

Private Sub lblO6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblO6)
End Sub

Private Sub lblP1_Click()
    objCaller.BackColor = lblP1.BackColor
    Unload Me
End Sub

Private Sub lblP1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblP1)
End Sub

Private Sub lblP2_Click()
    objCaller.BackColor = lblP2.BackColor
    Unload Me
End Sub

Private Sub lblP2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblP2)
End Sub

Private Sub lblP3_Click()
    objCaller.BackColor = lblP3.BackColor
    Unload Me
End Sub

Private Sub lblP3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblP3)
End Sub

Private Sub lblP4_Click()
    objCaller.BackColor = lblP4.BackColor
    Unload Me
End Sub

Private Sub lblP4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblP4)
End Sub

Private Sub lblP5_Click()
    objCaller.BackColor = lblP5.BackColor
    Unload Me
End Sub

Private Sub lblP5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblP5)
End Sub

Private Sub lblP6_Click()
    objCaller.BackColor = lblP6.BackColor
    Unload Me
End Sub

Private Sub lblP6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblP6)
End Sub

Private Sub lblR1_Click()
    objCaller.BackColor = lblR1.BackColor
    Unload Me
End Sub

Private Sub lblR1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblR1)
End Sub

Private Sub lblR2_Click()
    objCaller.BackColor = lblR2.BackColor
    Unload Me
End Sub

Private Sub lblR2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblR2)
End Sub

Private Sub lblR3_Click()
    objCaller.BackColor = lblR3.BackColor
    Unload Me
End Sub

Private Sub lblR3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblR3)
End Sub

Private Sub lblR4_Click()
    objCaller.BackColor = lblR4.BackColor
    Unload Me
End Sub

Private Sub lblR4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblR4)
End Sub

Private Sub lblR5_Click()
    objCaller.BackColor = lblR5.BackColor
    Unload Me
End Sub

Private Sub lblR5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblR5)
End Sub

Private Sub lblR6_Click()
    objCaller.BackColor = lblR6.BackColor
    Unload Me
End Sub

Private Sub lblR6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblR6)
End Sub

Private Sub lblT1_Click()
    objCaller.BackColor = lblT1.BackColor
    Unload Me
End Sub

Private Sub lblT1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblT1)
End Sub

Private Sub lblT2_Click()
    objCaller.BackColor = lblT2.BackColor
    Unload Me
End Sub

Private Sub lblT2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblT2)
End Sub

Private Sub lblT3_Click()
    objCaller.BackColor = lblT3.BackColor
    Unload Me
End Sub

Private Sub lblT3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblT3)
End Sub

Private Sub lblT4_Click()
    objCaller.BackColor = lblT4.BackColor
    Unload Me
End Sub

Private Sub lblT4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblT4)
End Sub

Private Sub lblT5_Click()
    objCaller.BackColor = lblT5.BackColor
    Unload Me
End Sub

Private Sub lblT5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblT5)
End Sub

Private Sub lblT6_Click()
    objCaller.BackColor = lblT6.BackColor
    Unload Me
End Sub

Private Sub lblT6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblT6)
End Sub

Private Sub lblW1_Click()
    objCaller.BackColor = lblW1.BackColor
    Unload Me
End Sub

Private Sub lblW1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblW1)
End Sub

Private Sub lblW2_Click()
    objCaller.BackColor = lblW2.BackColor
    Unload Me
End Sub

Private Sub lblW2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblW2)
End Sub

Private Sub lblW3_Click()
    objCaller.BackColor = lblW3.BackColor
    Unload Me
End Sub

Private Sub lblW3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblW3)
End Sub

Private Sub lblW4_Click()
    objCaller.BackColor = lblW4.BackColor
    Unload Me
End Sub

Private Sub lblW4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblW4)
End Sub

Private Sub lblW5_Click()
    objCaller.BackColor = lblW5.BackColor
    Unload Me
End Sub

Private Sub lblW5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblW5)
End Sub

Private Sub lblW6_Click()
    objCaller.BackColor = lblW6.BackColor
    Unload Me
End Sub

Private Sub lblW6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblW6)
End Sub

Private Sub lblY1_Click()
    objCaller.BackColor = lblY1.BackColor
    Unload Me
End Sub

Private Sub lblY1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblY1)
End Sub

Private Sub lblY2_Click()
    objCaller.BackColor = lblY2.BackColor
    Unload Me
End Sub

Private Sub lblY2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblY2)
End Sub

Private Sub lblY3_Click()
    objCaller.BackColor = lblY3.BackColor
    Unload Me
End Sub

Private Sub lblY3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblY3)
End Sub

Private Sub lblY4_Click()
    objCaller.BackColor = lblY4.BackColor
    Unload Me
End Sub

Private Sub lblY4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblY4)
End Sub

Private Sub lblY5_Click()
    objCaller.BackColor = lblY5.BackColor
    Unload Me
End Sub

Private Sub lblY5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblY5)
End Sub

Private Sub lblY6_Click()
    objCaller.BackColor = lblY6.BackColor
    Unload Me
End Sub

Private Sub lblY6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Call TurnOn(lblY6)
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call TurnOff(turnedOn)
    Set turnedOn = Nothing
End Sub

Private Sub UserForm_Terminate()
    Set objCaller = Nothing
End Sub
