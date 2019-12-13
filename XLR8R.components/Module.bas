Attribute VB_Name = "Module"
'Ini file path
Public INIPATH  As String

'ArrangeCursors
Public AC_SC    As String
Public AC_SHT   As String
Public AC_HOME  As String

'Highlighter
Public HL_BD    As String
Public HL_CO    As String

'SelectObjects
Public SO_SC    As String
Public SO_RNG   As String

'CopyAsBitmap
Public CB_SC    As String

'Language
Public LANG     As String

'Message
Public Const msgMultiEn As String = "Can't run this add-in on shared workbook."
Public Const msgMultiJp As String = "共有ワークブックでは実行できません｡"
Public Const msgMultiKr As String = "Can't run this add-in on shared workbook."
Public Const msgOlScEn  As String = "There are overlapping shortcuts."
Public Const msgOlScJp  As String = "重なるショートカットがあります。"
Public Const msgOlScKr  As String = "There are overlapping shortcuts."

Public Sub PreferenceSetting()
    If LANG = "jp" Then
        FormPS_jp.Show
    ElseIf LANG = "kr" Then
        FormPS_kr.Show
    Else
        FormPS.Show
    End If
End Sub

Public Sub ArrangeCursors()
    Dim clsAC As ClassAC
    Set clsAC = New ClassAC
    Call clsAC.pAC
    Set clsAC = Nothing
End Sub

Public Sub Highlighter_Border()
    Dim clsHL As ClassHL
    Set clsHL = New ClassHL
    Call clsHL.pHL("bd")
    Set clsHL = Nothing
End Sub

Public Sub Highlighter_Callout()
    Dim clsHL As ClassHL
    Set clsHL = New ClassHL
    Call clsHL.pHL("co")
    Set clsHL = Nothing
End Sub

Public Sub SelectObjects()
    Dim clsSO As ClassSO
    Set clsSO = New ClassSO
    Call clsSO.pSO
    Set clsSO = Nothing
End Sub

Public Sub CopyAsBitmap()
    Dim clsCB As ClassCB
    Set clsCB = New ClassCB
    Call clsCB.pCB
    Set clsCB = Nothing
End Sub

Public Sub SwitchSheet_First()
    Dim clsSS As ClassSS
    Set clsSS = New ClassSS
    Call clsSS.pSS_F
    Set clsSS = Nothing
End Sub

Public Sub SwitchSheet_Last()
    Dim clsSS As ClassSS
    Set clsSS = New ClassSS
    Call clsSS.pSS_L
    Set clsSS = Nothing
End Sub
