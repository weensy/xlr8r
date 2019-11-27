Attribute VB_Name = "Module"
'Ini file path
Public INIPATH  As String

'ArrangeCursors
Public AC_SC    As String
Public AC_SHT   As String
Public AC_HOME  As String

'Highlighter
Public HL_SC    As String
Public HL_SHP   As String

'SelectObjects
Public SO_SC    As String
Public SO_RNG   As String

'CopyAsBitmap
Public CB_SC    As String

Public Sub PreferenceSetting()
    FormPS.Show
End Sub

Public Sub ArrangeCursors()
    Dim clsAC As ClassAC
    Set clsAC = New ClassAC
    Call clsAC.pAC
    Set clsAC = Nothing
End Sub

Public Sub Highlighter()
    Dim clsHL As ClassHL
    Set clsHL = New ClassHL
    Call clsHL.pHL
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
