Attribute VB_Name = "Module"
Public Const CURRENT_VERSION As String = "v1.2.5"
Public LATEST_VERSION As String

'Ini file path
Public INIPATH          As String

'ArrangeCursors
Public AC_SC            As String
Public AC_SHT           As String
Public AC_HOME          As String

'Highlighter
Public HL_OL            As String
Public HL_CO            As String
Public HL_OL_CLR_LINE   As Long
Public HL_CO_CLR_LINE   As Long
Public HL_CO_CLR_FONT   As Long

'SelectObjects
Public SO_SC            As String
Public SO_RNG           As String

'CopyAsBitmap
Public CB_SC            As String

'Language
Public LANG             As String

'Message
Public msgMulti         As String
Public msgOlSc          As String
Public msgCdStt         As String
Public msgCdEnd         As String
Public msgNeCll         As String
Public msgExRow         As String
Public msgExCol         As String
Public ttlCdStt         As String
Public ttlCdEnd         As String

'Color Picker
Public objCaller        As Object

Public Sub PreferenceSetting()
    FormPS.Show
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
    Call clsHL.pHL("ol")
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
