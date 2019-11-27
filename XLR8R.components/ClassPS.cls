VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function WriteString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AN$, ByVal KN$, ByVal iVal$, ByVal fPath$)
Private Declare Function GetString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal AN$, ByVal KN$, ByVal iVal$, ByVal RetS$, ByVal rSize&, ByVal fPath$)
 
Public Sub WriteINI(iSt As String, iK As String, iVal As String)
    Call WriteString(iSt, iK, iVal, INIPATH)
End Sub
 
Public Function ReadINI(iSt As String, iK As String) As String
    Dim sBuf As String
    Dim sSize As Long
    sBuf = Space$(255)
    sSize = GetString(iSt, iK, "", sBuf, 255, INIPATH)
    If sSize Then ReadINI = Left$(sBuf, sSize)
End Function

Public Sub SetSC(sKey As String, sProc As String)
    If sKey = "" Then
        Exit Sub
    End If
    
    If sProc = "" Then
        Application.OnKey "^+" & sKey
    Else
        Application.OnKey "^+" & sKey, sProc
    End If
End Sub
