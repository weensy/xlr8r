VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormPS 
   Caption         =   "Preference Settings"
   ClientHeight    =   3480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   OleObjectBlob   =   "FormPS.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "FormPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnDefault_Click()
    'ArrangeCursors
    chkAcSc.Value = True
    txtAcSc.Value = "A"
    optAcCs.Value = True
    optAcA1.Value = True
    
    'Highlighter
    chkHlBd.Value = True
    txtHlBd.Value = "D"
    chkHlCo.Value = True
    txtHlCo.Value = "F"
    
    'SelectObjects
    chkSoSc.Value = True
    txtSoSc.Value = "S"
    optSoSl.Value = True
    
    'CopyAsBitmap
    chkCbSc.Value = True
    txtCbSc.Value = "C"
    
    'Language
    optEnglish.Value = True
    
End Sub

Private Sub btnOK_Click()

    'Check
    Dim sc(4) As String
    
    sc(0) = txtAcSc.Value
    sc(1) = txtHlBd.Value
    sc(2) = txtHlCo.Value
    sc(3) = txtSoSc.Value
    sc(4) = txtCbSc.Value
    
    For i = 0 To 3
        For j = i + 1 To 4
            If sc(i) = sc(j) Then
                If LANG = "jp" Then
                    MsgBox msgOlScJp
                ElseIf LANG = "kr" Then
                    MsgBox msgOlScKr
                Else
                    MsgBox msgOlScEn
                End If
                Exit Sub
            End If
        Next
    Next

    'Work
    Dim ClsPS As ClassPS
    Set ClsPS = New ClassPS
    
    Call ClsPS.SetSC(AC_SC, "")
    Call ClsPS.SetSC(HL_BD, "")
    Call ClsPS.SetSC(HL_CO, "")
    Call ClsPS.SetSC(SO_SC, "")
    Call ClsPS.SetSC(CB_SC, "")
    
    AC_SC = txtAcSc.Value
    HL_BD = txtHlBd.Value
    HL_CO = txtHlCo.Value
    SO_SC = txtSoSc.Value
    CB_SC = txtCbSc.Value

    Call ClsPS.SetSC(AC_SC, "ArrangeCursors")
    Call ClsPS.SetSC(HL_BD, "Highlighter_Border")
    Call ClsPS.SetSC(HL_CO, "Highlighter_Callout")
    Call ClsPS.SetSC(SO_SC, "SelectObjects")
    Call ClsPS.SetSC(CB_SC, "CopyAsBitmap")
    
    If optAcFs.Value Then
        AC_SHT = "fs"
    Else
        AC_SHT = "cs"
    End If
    
    AC_HOME = txtAcCstm.Value
        
    If optSoCd.Value Then
        SO_RNG = "cd"
    Else
        SO_RNG = "sl"
    End If
    
    If optJapanese.Value Then
        LANG = "jp"
    ElseIf optKorean.Value Then
        LANG = "kr"
    Else
        LANG = "en"
    End If
    
    Call ClsPS.WriteINI("ArrangeCursors", "AC_SC", AC_SC)
    Call ClsPS.WriteINI("ArrangeCursors", "AC_SHT", AC_SHT)
    Call ClsPS.WriteINI("ArrangeCursors", "AC_HOME", AC_HOME)
    Call ClsPS.WriteINI("Highlighter", "HL_BD", HL_BD)
    Call ClsPS.WriteINI("Highlighter", "HL_CO", HL_CO)
    Call ClsPS.WriteINI("SelectObjects", "SO_SC", SO_SC)
    Call ClsPS.WriteINI("SelectObjects", "SO_RNG", SO_RNG)
    Call ClsPS.WriteINI("CopyAsBitmap", "CB_SC", CB_SC)
    Call ClsPS.WriteINI("Language", "LANG", LANG)

    Set ClsPS = Nothing
    
    Unload Me
End Sub

Private Sub chkAcSc_Change()
    If chkAcSc.Value Then
        txtAcSc.Enabled = True
        If MultiPage.Value = 0 Then
            txtAcSc.SetFocus
        End If
    Else
        txtAcSc.Enabled = False
        txtAcSc.Value = ""
    End If
End Sub

Private Sub chkCbSc_Change()
    If chkCbSc.Value Then
        txtCbSc.Enabled = True
        If MultiPage.Value = 3 Then
            txtCbSc.SetFocus
        End If
    Else
        txtCbSc.Enabled = False
        txtCbSc.Value = ""
    End If
End Sub

Private Sub chkHlBd_Change()
    If chkHlBd.Value Then
        txtHlBd.Enabled = True
        If MultiPage.Value = 1 Then
            txtHlBd.SetFocus
        End If
    Else
        txtHlBd.Enabled = False
        txtHlBd.Value = ""
    End If
End Sub

Private Sub chkHlCo_Change()
    If chkHlCo.Value Then
        txtHlCo.Enabled = True
        If MultiPage.Value = 1 Then
            txtHlCo.SetFocus
        End If
    Else
        txtHlCo.Enabled = False
        txtHlCo.Value = ""
    End If
End Sub

Private Sub chkSoSc_Change()
    If chkSoSc.Value Then
        txtSoSc.Enabled = True
        If MultiPage.Value = 2 Then
            txtSoSc.SetFocus
        End If
    Else
        txtSoSc.Enabled = False
        txtSoSc.Value = ""
    End If
End Sub

Private Sub optAcCstm_Change()
    If optAcCstm.Value Then
        txtAcCstm.Enabled = True
        If MultiPage.Value = 0 Then
            txtAcCstm.SetFocus
        End If
    Else
        txtAcCstm.Enabled = False
        txtAcCstm.Value = ""
    End If
End Sub

Private Sub txtAcCstm_Change()
    txtAcCstm.Value = UCase(txtAcCstm.Value)
End Sub

Private Sub txtAcSc_Change()
    txtAcSc.Value = UCase(txtAcSc.Value)
End Sub

Private Sub txtCbSc_Change()
    txtCbSc.Value = UCase(txtCbSc.Value)
End Sub

Private Sub txtHlSc_Change()
    txtHlSc.Value = UCase(txtHlSc.Value)
End Sub

Private Sub txtSoSc_Change()
    txtSoSc.Value = UCase(txtSoSc.Value)
End Sub

Private Sub UserForm_Initialize()
    '********************************
    'ArrangeCursor
    '********************************
    If AC_SC = "" Then
        chkAcSc.Value = False
        txtAcSc.Enabled = False
    Else
        chkAcSc.Value = True
    End If
    txtAcSc.Value = AC_SC
    If AC_SHT = "fs" Then
        optAcFs.Value = True
        optAcCs.Value = False
    Else
        optAcFs.Value = False
        optAcCs.Value = True
    End If
    If AC_HOME = "" Then
        optAcA1.Value = True
        optAcCstm.Value = False
        txtAcCstm.Enabled = False
    Else
        optAcA1.Value = False
        optAcCstm.Value = True
    End If
    txtAcCstm.Value = AC_HOME
    
    '********************************
    'Highligher
    '********************************
    If HL_BD = "" Then
        chkHlBd.Value = False
        txtHlBd.Enabled = False
    Else
        chkHlBd.Value = True
    End If
    txtHlBd.Value = HL_BD
    If HL_CO = "" Then
        chkHlCo.Value = False
        txtHlCo.Enabled = False
    Else
        chkHlCo.Value = True
    End If
    txtHlCo.Value = HL_CO
    
    '********************************
    'SelectObjects
    '********************************
    If SO_SC = "" Then
        chkSoSc.Value = False
        txtSoSc.Enabled = False
    Else
        chkSoSc.Value = True
    End If
    txtSoSc.Value = SO_SC
    If SO_RNG = "cd" Then
        optSoSl.Value = False
        optSoCd.Value = True
    Else
        optSoSl.Value = True
        optSoCd.Value = False
    End If
    
    '********************************
    'CopyAsBitmap
    '********************************
    If CB_SC = "" Then
        chkCbSc.Value = False
        txtCbSc.Enabled = False
    Else
        chkCbSc.Value = True
    End If
    txtCbSc.Value = CB_SC
    
    '********************************
    'Language
    '********************************
    If LANG = "jp" Then
        optEnglish.Value = False
        optJapanese.Value = True
        optKorean.Value = False
    ElseIf LANG = "kr" Then
        optEnglish.Value = False
        optJapanese.Value = False
        optKorean.Value = True
    Else
        optEnglish.Value = True
        optJapanese.Value = False
        optKorean.Value = False
    End If
End Sub
