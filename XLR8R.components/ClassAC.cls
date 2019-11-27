VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub pAC()
    Dim i               As Long
    Dim j               As Long
    Dim frzpnsHomeRow   As Long
    Dim frzpnsHomeCol   As Long
    Dim activated       As Long
    
    'Save index of activated sheet
    activated = ActiveSheet.Index
    
    'Set cursor
    For i = Sheets.Count To 1 Step -1
        If Sheets(i).Visible = xlSheetVisible Then
            Sheets(i).Select
            If ActiveWindow.FreezePanes Then
                frzpnsHomeRow = 0
                frzpnsHomeCol = 0
                For j = 1 To Rows.Count - ActiveWindow.SplitRow
                    If Not Rows(ActiveWindow.SplitRow + j).Hidden Then
                        frzpnsHomeRow = ActiveWindow.SplitRow + j
                        Exit For
                    End If
                Next j
                For j = 1 To Columns.Count - ActiveWindow.SplitColumn
                    If Not Columns(ActiveWindow.SplitColumn + j).Hidden Then
                        frzpnsHomeCol = ActiveWindow.SplitColumn + j
                        Exit For
                    End If
                Next j
                Sheets(i).Cells(frzpnsHomeRow, frzpnsHomeCol).Select
            End If
            
            If AC_HOME = "" Then
                Sheets(i).Cells(1, 1).Select
            Else
                Sheets(i).Range(AC_HOME).Select
            End If
        End If
    Next i
    
    'Active sheet
    If AC_SHT <> "fs" Then
        Sheets(activated).Select
    End If
End Sub
