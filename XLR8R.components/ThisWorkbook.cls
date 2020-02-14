VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If ThisWorkbook.VBProject.Saved Then
        Exit Sub
    End If
    
    Dim expPath As String
    Dim expName As String
    
    expPath = Left(ThisWorkbook.FullName, InStr(ThisWorkbook.FullName, ".xlam")) & "components"
    
    If Dir(expPath, vbDirectory) = "" Then
        MkDir expPath
    Else
        buf = Dir(expPath & "\")
        Do While buf <> ""
            Kill expPath & "\" & buf
            buf = Dir()
        Loop
    End If
    
    For Each comp In ThisWorkbook.VBProject.VBComponents
        expName = comp.Name
        
        If comp.Type = 1 Then
            expName = expName & ".bas"
        ElseIf comp.Type = 3 Then
            expName = expName & ".frm"
        ElseIf comp.Type = 2 Or comp.Type = 100 Then
            expName = expName & ".cls"
        End If
        
        comp.Export expPath & "\" & expName
    Next
End Sub

Private Sub Workbook_Open()
    INIPATH = ThisWorkbook.Path & "\XLR8R.ini"
    
    Dim ClsPS As ClassPS
    Set ClsPS = New ClassPS
    
    'ArrangeCursors
    AC_SC = ClsPS.ReadINI("ArrangeCursors", "AC_SC")
    AC_SHT = ClsPS.ReadINI("ArrangeCursors", "AC_SHT")
    AC_HOME = ClsPS.ReadINI("ArrangeCursors", "AC_HOME")
    
    'Highlighter
    HL_OL = ClsPS.ReadINI("Highlighter", "HL_OL")
    HL_CO = ClsPS.ReadINI("Highlighter", "HL_CO")
    HL_OL_CLR_LINE = ClsPS.ReadINI4Long("Highlighter", "HL_OL_CLR_LINE")
    HL_CO_CLR_LINE = ClsPS.ReadINI4Long("Highlighter", "HL_CO_CLR_LINE")
    HL_CO_CLR_FONT = ClsPS.ReadINI4Long("Highlighter", "HL_CO_CLR_FONT")
    
    'SelectObjects
    SO_SC = ClsPS.ReadINI("SelectObjects", "SO_SC")
    SO_RNG = ClsPS.ReadINI("SelectObjects", "SO_RNG")
    
    'CopyAsBitmap
    CB_SC = ClsPS.ReadINI("CopyAsBitmap", "CB_SC")
    
    'Language
    LANG = ClsPS.ReadINI("Language", "LANG")
    
    'Message
    Dim ClsCL As ClassCL
    Set ClsCL = New ClassCL
    Call ClsCL.SetMsg
    Set ClsCL = Nothing
    
    'Set Shortcut
    Call ClsPS.SetSC("p", "PreferenceSetting")
    Call ClsPS.SetSC(AC_SC, "ArrangeCursors")
    Call ClsPS.SetSC(HL_OL, "Highlighter_Border")
    Call ClsPS.SetSC(HL_CO, "Highlighter_Callout")
    Call ClsPS.SetSC(SO_SC, "SelectObjects")
    Call ClsPS.SetSC(CB_SC, "CopyAsBitmap")

    Application.OnKey "^%{HOME}", "SwitchSheet_First"
    Application.OnKey "^%{END}", "SwitchSheet_Last"
    
    Set ClsPS = Nothing
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim ClsPS As ClassPS
    Set ClsPS = New ClassPS
    
    'Reset Shortcut
    Call ClsPS.SetSC("p", "")
    Call ClsPS.SetSC(AC_SC, "")
    Call ClsPS.SetSC(HL_OL, "")
    Call ClsPS.SetSC(HL_CO, "")
    Call ClsPS.SetSC(SO_SC, "")
    Call ClsPS.SetSC(CB_SC, "")
    
    Application.OnKey "^%{HOME}"
    Application.OnKey "^%{END}"
    
    Set ClsPS = Nothing
End Sub
