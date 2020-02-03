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

Private Sub btnChkUpd_Click()

    Dim WinHttp   As Object
    Dim Source    As String
    Dim buf       As Long
    
    Set WinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    WinHttp.Open "GET", "https://github.com/vaporwavy/xlr8r/releases.atom"
    WinHttp.Send
    WinHttp.WaitForResponse
    Source = WinHttp.ResponseText
    
    buf = InStr(InStr(Source, "<entry>"), Source, "<title>") + 7
    LATEST_VERSION = Mid(Source, buf, InStr(buf, Source, "</title>") - buf)
    
    lblLatVer.Caption = LATEST_VERSION
    btnChkUpd.Visible = False
    btnDownload.Visible = True
    
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
End Sub

Private Sub btnDownload_Click()
    ThisWorkbook.FollowHyperlink "https://github.com/vaporwavy/xlr8r/archive/" & LATEST_VERSION & ".zip"
End Sub

Private Sub btnOK_Click()

    'Check shortcut
    Dim sc(4) As String
    
    sc(0) = txtAcSc.Value
    sc(1) = txtHlBd.Value
    sc(2) = txtHlCo.Value
    sc(3) = txtSoSc.Value
    sc(4) = txtCbSc.Value
    
    For i = 0 To 3
        For j = i + 1 To 4
            If sc(i) <> "" _
            And sc(j) <> "" _
            And sc(i) = sc(j) Then
                MsgBox msgOlSc
                Exit Sub
            End If
        Next
    Next
    
    'Check cell
    If optAcCstm.Value Then
    
        Dim RE      As Object
        Dim mCol    As Object
        Dim mRow    As Object
        Dim nCol    As Long
        Dim nRow    As Long
        Dim flgOver As Boolean
        
        Set RE = CreateObject("VBScript.RegExp")
        RE.Pattern = "^[A-Z]+[0-9]+$"
        
        If Not RE.Test(txtAcCstm.Value) Then
            MsgBox msgNeCll
            MultiPage.Value = 1
            txtAcCstm.SetFocus
            Exit Sub
        End If
        
        RE.Pattern = "[A-Z]+"
        
        Set mCol = RE.Execute(txtAcCstm.Value)
        
        RE.Pattern = "[0-9]+"
        
        Set mRow = RE.Execute(txtAcCstm.Value)
        
        Set RE = Nothing
        
        If mCol(0).Length > 7 Then
            flgOver = True
        Else
            flgOver = False
        
            nCol = C2N(Mid(mCol(0).Value, mCol(0).Length, 1))
            
            For i = 1 To mCol(0).Length - 1
                nCol = nCol + 26 ^ (mCol(0).Length - i) * C2N(Mid(mCol(0).Value, i, 1))
            Next
        End If
        
        Set mCol = Nothing
        
        nRow = mRow(0).Value
        
        Set mRow = Nothing
        
        If ActiveWorkbook.ActiveSheet.Columns.Count < nCol _
        Or flgOver Then
            MsgBox msgExCol
            MultiPage.Value = 1
            txtAcCstm.SetFocus
            Exit Sub
        ElseIf ActiveWorkbook.ActiveSheet.Rows.Count < nRow Then
            MsgBox msgExRow
            MultiPage.Value = 1
            txtAcCstm.SetFocus
            Exit Sub
        End If
        
    End If

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
    Else
        LANG = "en"
    End If
    
    Dim ClsCL As ClassCL
    Set ClsCL = New ClassCL
    Call ClsCL.SetMsg
    Set ClsCL = Nothing
    
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
        If MultiPage.Value = 1 Then
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
        If MultiPage.Value = 4 Then
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
        If MultiPage.Value = 2 Then
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
        If MultiPage.Value = 2 Then
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
        If MultiPage.Value = 3 Then
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
        If MultiPage.Value = 1 Then
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
    'General
    '********************************
    lblCurVer.Caption = CURRENT_VERSION
    lblLatVer.Caption = LATEST_VERSION
    If LATEST_VERSION = "" Then
        btnChkUpd.Visible = True
        btnDownload.Visible = False
    Else
        btnChkUpd.Visible = False
        btnDownload.Visible = True
    End If
    
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
    Else
        optEnglish.Value = True
        optJapanese.Value = False
    End If
    
    '********************************
    'Set caption
    '********************************
    If LANG = "jp" Then
        FormPS.Caption = "設定"
        btnDefault.Caption = "デフォルト"
        btnCancel.Caption = "キャンセル"
        MultiPage.Pages("pgAC").Caption = "カーソル配置"
        frmAcSht.Caption = "シートの位置"
        optAcFs.Caption = "最初シート"
        optAcCs.Caption = "現在シート"
        frmAcCl.Caption = "セルの位置"
        optAcCstm.Caption = "指定"
        MultiPage.Pages("pgHL").Caption = "ハイライト"
        frmHlBd.Caption = "枠"
        frmHlCo.Caption = "吹き出し"
        MultiPage.Pages("pgSO").Caption = "オブジェクト選択"
        frmSoRng.Caption = "範囲"
        optSoSl.Caption = "選択範囲"
        optSoCd.Caption = "座標利用"
        MultiPage.Pages("pgCB").Caption = "イメージでコピー"
        MultiPage.Pages("pgSS").Caption = "シート切り替え"
        frmSF.Caption = "最初シートへ"
        frmSL.Caption = "最後シートへ"
        MultiPage.Pages("pgGS").Caption = "全般"
        frmLng.Caption = "言語"
        optEnglish.Caption = "英語"
        optJapanese.Caption = "日本語"
        frmVer.Caption = "バージョン"
        lblCurrent.Caption = "現在"
        lblLatest.Caption = "最新"
        btnChkUpd.Caption = "チェック"
        btnDownload.Caption = "ダウンロード"
    End If
    
End Sub

Private Function C2N(str As String)
    If str = "A" Then
        C2N = 1
    ElseIf str = "B" Then
        C2N = 2
    ElseIf str = "C" Then
        C2N = 3
    ElseIf str = "D" Then
        C2N = 4
    ElseIf str = "E" Then
        C2N = 5
    ElseIf str = "F" Then
        C2N = 6
    ElseIf str = "G" Then
        C2N = 7
    ElseIf str = "H" Then
        C2N = 8
    ElseIf str = "I" Then
        C2N = 9
    ElseIf str = "J" Then
        C2N = 10
    ElseIf str = "K" Then
        C2N = 11
    ElseIf str = "L" Then
        C2N = 12
    ElseIf str = "M" Then
        C2N = 13
    ElseIf str = "N" Then
        C2N = 14
    ElseIf str = "O" Then
        C2N = 15
    ElseIf str = "P" Then
        C2N = 16
    ElseIf str = "Q" Then
        C2N = 17
    ElseIf str = "R" Then
        C2N = 18
    ElseIf str = "S" Then
        C2N = 19
    ElseIf str = "T" Then
        C2N = 20
    ElseIf str = "U" Then
        C2N = 21
    ElseIf str = "V" Then
        C2N = 22
    ElseIf str = "W" Then
        C2N = 23
    ElseIf str = "X" Then
        C2N = 24
    ElseIf str = "Y" Then
        C2N = 25
    ElseIf str = "Z" Then
        C2N = 26
    Else
        C2N = 0
    End If
End Function
