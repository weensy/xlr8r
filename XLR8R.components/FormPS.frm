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
Private Const URL_RELEASE As String = "https://github.com/vaporwavy/xlr8r/releases.atom"
Private Const URL_ARCHIVE As String = "https://github.com/vaporwavy/xlr8r/archive/"

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnChkUpd_Click()

    Dim WinHttp   As Object
    Dim Source    As String
    Dim pStart    As Long
    
    Set WinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    WinHttp.Open "GET", URL_RELEASE
    WinHttp.Send
    WinHttp.WaitForResponse
    Source = WinHttp.ResponseText
    
    pStart = InStr(InStr(Source, "<entry>"), Source, "<title>") + 7
    LATEST_VERSION = Mid(Source, pStart, InStr(pStart, Source, "</title>") - pStart)
    
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
    chkHlOl.Value = True
    txtHlOl.Value = "D"
    chkHlCo.Value = True
    txtHlCo.Value = "F"
    lblHlOlClrValLine.BackColor = 255
    lblHlCoClrValLine.BackColor = 255
    lblHlCoClrValFont.BackColor = 255
    
    'SelectObjects
    chkSoSc.Value = True
    txtSoSc.Value = "S"
    optSoSl.Value = True
    
    'CopyAsBitmap
    chkCbSc.Value = True
    txtCbSc.Value = "C"
End Sub

Private Sub btnDownload_Click()
    ThisWorkbook.FollowHyperlink URL_ARCHIVE & LATEST_VERSION & ".zip"
End Sub

Private Sub btnOK_Click()

    'Check shortcut
    Dim sc(4) As String
    
    sc(0) = txtAcSc.Value
    sc(1) = txtHlOl.Value
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
        
        If Not RE.test(txtAcCstm.Value) Then
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
    Call ClsPS.SetSC(HL_OL, "")
    Call ClsPS.SetSC(HL_CO, "")
    Call ClsPS.SetSC(SO_SC, "")
    Call ClsPS.SetSC(CB_SC, "")
    
    AC_SC = txtAcSc.Value
    HL_OL = txtHlOl.Value
    HL_CO = txtHlCo.Value
    SO_SC = txtSoSc.Value
    CB_SC = txtCbSc.Value

    Call ClsPS.SetSC(AC_SC, "ArrangeCursors")
    Call ClsPS.SetSC(HL_OL, "Highlighter_Border")
    Call ClsPS.SetSC(HL_CO, "Highlighter_Callout")
    Call ClsPS.SetSC(SO_SC, "SelectObjects")
    Call ClsPS.SetSC(CB_SC, "CopyAsBitmap")
    
    If optAcFs.Value Then
        AC_SHT = "fs"
    Else
        AC_SHT = "cs"
    End If
    
    AC_HOME = txtAcCstm.Value
        
    HL_OL_CLR_LINE = lblHlOlClrValLine.BackColor
    HL_CO_CLR_LINE = lblHlCoClrValLine.BackColor
    HL_CO_CLR_FONT = lblHlCoClrValFont.BackColor
        
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
    Call ClsPS.WriteINI("Highlighter", "HL_OL", HL_OL)
    Call ClsPS.WriteINI("Highlighter", "HL_CO", HL_CO)
    Call ClsPS.WriteINI4Long("Highlighter", "HL_OL_CLR_LINE", HL_OL_CLR_LINE)
    Call ClsPS.WriteINI4Long("Highlighter", "HL_CO_CLR_LINE", HL_CO_CLR_LINE)
    Call ClsPS.WriteINI4Long("Highlighter", "HL_CO_CLR_FONT", HL_CO_CLR_FONT)
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

Private Sub chkHlOl_Change()
    If chkHlOl.Value Then
        txtHlOl.Enabled = True
        If MultiPage.Value = 2 Then
            txtHlOl.SetFocus
        End If
    Else
        txtHlOl.Enabled = False
        txtHlOl.Value = ""
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

Private Sub lblHlCoClrValFont_Click()
    Set objCaller = lblHlCoClrValFont
    FormCP.Show
End Sub

Private Sub lblHlCoClrValLine_Click()
    Set objCaller = lblHlCoClrValLine
    FormCP.Show
End Sub

Private Sub lblHlOlClrValLine_Click()
    Set objCaller = lblHlOlClrValLine
    FormCP.Show
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
    If HL_OL = "" Then
        chkHlOl.Value = False
        txtHlOl.Enabled = False
    Else
        chkHlOl.Value = True
    End If
    txtHlOl.Value = HL_OL
    If HL_CO = "" Then
        chkHlCo.Value = False
        txtHlCo.Enabled = False
    Else
        chkHlCo.Value = True
    End If
    txtHlCo.Value = HL_CO
    
    lblHlOlClrValLine.BackColor = HL_OL_CLR_LINE
    lblHlCoClrValLine.BackColor = HL_CO_CLR_LINE
    lblHlCoClrValFont.BackColor = HL_CO_CLR_FONT
    
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
        btnDefault.Caption = "デフォルト設定"
        btnCancel.Caption = "キャンセル"
        MultiPage.Pages("pgAC").Caption = "カーソル配置"
        frmAcSht.Caption = "シートの位置"
        optAcFs.Caption = "最初シート"
        optAcCs.Caption = "現在シート"
        frmAcCl.Caption = "セルの位置"
        optAcCstm.Caption = "指定"
        MultiPage.Pages("pgHL").Caption = "ハイライト"
        frmHlOl.Caption = "枠"
        frmHlCo.Caption = "吹き出し"
        lblHlOlClrLine.Caption = "ライン"
        lblHlCoClrLine.Caption = "ライン"
        lblHlCoClrFont.Caption = "フォント"
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

Private Function C2N(col As String)
    If col = "A" Then
        C2N = 1
    ElseIf col = "B" Then
        C2N = 2
    ElseIf col = "C" Then
        C2N = 3
    ElseIf col = "D" Then
        C2N = 4
    ElseIf col = "E" Then
        C2N = 5
    ElseIf col = "F" Then
        C2N = 6
    ElseIf col = "G" Then
        C2N = 7
    ElseIf col = "H" Then
        C2N = 8
    ElseIf col = "I" Then
        C2N = 9
    ElseIf col = "J" Then
        C2N = 10
    ElseIf col = "K" Then
        C2N = 11
    ElseIf col = "L" Then
        C2N = 12
    ElseIf col = "M" Then
        C2N = 13
    ElseIf col = "N" Then
        C2N = 14
    ElseIf col = "O" Then
        C2N = 15
    ElseIf col = "P" Then
        C2N = 16
    ElseIf col = "Q" Then
        C2N = 17
    ElseIf col = "R" Then
        C2N = 18
    ElseIf col = "S" Then
        C2N = 19
    ElseIf col = "T" Then
        C2N = 20
    ElseIf col = "U" Then
        C2N = 21
    ElseIf col = "V" Then
        C2N = 22
    ElseIf col = "W" Then
        C2N = 23
    ElseIf col = "X" Then
        C2N = 24
    ElseIf col = "Y" Then
        C2N = 25
    ElseIf col = "Z" Then
        C2N = 26
    Else
        C2N = 0
    End If
End Function

