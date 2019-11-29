VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassSO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetCursorPos Lib "USER32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const DPI As Long = 96
Private Const PPI As Long = 72

Public Sub pSO()
    If SO_RNG = "cd" Then
        Call SoCd
    Else
        Call SoSl
    End If
End Sub

Private Sub SoCd()
    
    Dim R1C1Left    As Double
    Dim R1C1Top     As Double
    Dim Pos         As POINTAPI
    Dim MouseL1     As Double
    Dim MouseT1     As Double
    Dim MouseL2     As Double
    Dim MouseT2     As Double
    Dim x1          As Double
    Dim x2          As Double
    Dim y1          As Double
    Dim y2          As Double
    Dim obj         As Object
    Dim zoomed      As Long
    Dim scrolledCol As Long
    Dim scrolledRow As Long
    
    If ActiveWorkbook.MultiUserEditing Then
        MsgBox "Can't run this add-in on shared workbook."
        Exit Sub
    End If
    
    Application.Cursor = xlNorthwestArrow
    
    'Zoom (To avoid excel bug)
    zoomed = ActiveWindow.Zoom
    scrolledCol = ActiveWindow.ScrollColumn
    scrolledRow = ActiveWindow.ScrollRow
    ActiveWindow.Zoom = 100
    ActiveWindow.ScrollColumn = scrolledCol
    ActiveWindow.ScrollRow = scrolledRow
    
    'Messagebox
    If MsgBox("Press the spacebar where you want.", vbOKCancel, "Start point") <> vbOK Then
        GoTo SetBefore
    End If
    
    'Get coordinate1
    R1C1Left = ActiveWindow.ActivePane.PointsToScreenPixelsX(0) '(Cells(scrolledRow, scrolledCol).Left)
    R1C1Top = ActiveWindow.ActivePane.PointsToScreenPixelsY(0) '(Cells(scrolledRow, scrolledCol).Top)
    GetCursorPos Pos
    MouseL1 = (Pos.x - R1C1Left) / (DPI / PPI) '/ (ActiveWindow.Zoom / 100)
    MouseT1 = (Pos.y - R1C1Top) / (DPI / PPI) '/ (ActiveWindow.Zoom / 100)
    
    'Messagebox
    If MsgBox("Press the spacebar where you want.", vbOKCancel, "End point") <> vbOK Then
        GoTo SetBefore
    End If
    
    'Get coordinate2
    R1C1Left = ActiveWindow.ActivePane.PointsToScreenPixelsX(0) '(Cells(scrolledRow, scrolledCol).Left)
    R1C1Top = ActiveWindow.ActivePane.PointsToScreenPixelsY(0) '(Cells(scrolledRow, scrolledCol).Top)
    GetCursorPos Pos
    MouseL2 = (Pos.x - R1C1Left) / (DPI / PPI) '/ (ActiveWindow.Zoom / 100)
    MouseT2 = (Pos.y - R1C1Top) / (DPI / PPI) '/ (ActiveWindow.Zoom / 100)
    
    'Analysis
    If MouseL1 < MouseL2 Then
        x1 = MouseL1 '+ Cells(scrolledRow, scrolledCol).Left
        x2 = MouseL2 '+ Cells(scrolledRow, scrolledCol).Left
    Else
        x1 = MouseL2 '+ Cells(scrolledRow, scrolledCol).Left
        x2 = MouseL1 '+ Cells(scrolledRow, scrolledCol).Left
    End If
    If MouseT1 < MouseT2 Then
        y1 = MouseT1 '+ Cells(scrolledRow, scrolledCol).Top
        y2 = MouseT2 '+ Cells(scrolledRow, scrolledCol).Top
    Else
        y1 = MouseT2 '+ Cells(scrolledRow, scrolledCol).Top
        y2 = MouseT1 '+ Cells(scrolledRow, scrolledCol).Top
    End If
    
    Application.ScreenUpdating = False
    
    'Select object
    For Each obj In ActiveSheet.Shapes
        If obj.Left >= x1 And _
        obj.Top >= y1 And _
        obj.Left + obj.Width <= x2 And _
        obj.Top + obj.Height <= y2 And _
        obj.Visible = True Then
            obj.Select Replace:=False
        End If
    Next
    
    Application.ScreenUpdating = True

SetBefore:
    'Rezoom (To avoid excel bug)
    ActiveWindow.Zoom = zoomed
    ActiveWindow.ScrollColumn = scrolledCol
    ActiveWindow.ScrollRow = scrolledRow
    
    Application.Cursor = xlDefault

End Sub

Private Sub SoSl()

    If TypeName(Selection) <> "Range" Then
        If Selection.ShapeRange.Count > 1 Then
            MsgBox "Can not work when more than one object is selected."
            Exit Sub
        End If
    End If
    
    If ActiveWorkbook.MultiUserEditing Then
        MsgBox "Can't run this add-in on shared workbook."
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    Dim x1 As Double
    Dim y1 As Double
    Dim x2 As Double
    Dim y2 As Double
    Dim obj As Object
    Dim zoomed As Long
    
    'Zoom (To avoid excel bug)
    zoomed = ActiveWindow.Zoom
    ActiveWindow.Zoom = 100
    
    'Get coordinate
    x1 = Selection.Left
    y1 = Selection.Top
    x2 = x1 + Selection.Width
    y2 = y1 + Selection.Height
    
    'Select object
    For Each obj In ActiveSheet.Shapes
        If obj.Left >= x1 And _
        obj.Top >= y1 And _
        obj.Left + obj.Width <= x2 And _
        obj.Top + obj.Height <= y2 And _
        obj.Visible = True Then
            obj.Select Replace:=False
        End If
    Next

    'Rezoom (To avoid excel bug)
    ActiveWindow.Zoom = zoomed

    Application.ScreenUpdating = True
    
End Sub