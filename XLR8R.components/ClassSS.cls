VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub pSS_F()
    For i = 1 To ActiveWorkbook.Sheets.Count
        If Sheets(i).Visible = xlSheetVisible Then
            'Sheets(i).Activate
            Sheets(i).Select
            Exit For
        End If
    Next
End Sub

Public Sub pSS_L()
    For i = ActiveWorkbook.Sheets.Count To 1 Step -1
        If Sheets(i).Visible = xlSheetVisible Then
            'Sheets(i).Activate
            Sheets(i).Select
            Exit For
        End If
    Next
End Sub
