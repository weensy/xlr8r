VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassHL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare PtrSafe Function GetCursorPos Lib "USER32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const DPI As Long = 96
Private Const PPI As Long = 72

Public Sub pHL(HL_SHP As String)
    
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
    Dim shp         As Shape
    Dim zoomed      As Long
    Dim scrolledCol As Long
    Dim scrolledRow As Long
        
    If ActiveWorkbook.MultiUserEditing Then
        MsgBox msgMulti
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
    If MsgBox(msgCdStt, vbOKCancel, ttlCdStt) <> vbOK Then
        GoTo SetBefore
    End If
    
    'Get coordinate1
    R1C1Left = ActiveWindow.ActivePane.PointsToScreenPixelsX(0) '(Cells(scrolledRow, scrolledCol).Left)
    R1C1Top = ActiveWindow.ActivePane.PointsToScreenPixelsY(0) '(Cells(scrolledRow, scrolledCol).Top)
    GetCursorPos Pos
    MouseL1 = (Pos.x - R1C1Left) / (DPI / PPI) '/ (ActiveWindow.Zoom / 100)
    MouseT1 = (Pos.y - R1C1Top) / (DPI / PPI) '/ (ActiveWindow.Zoom / 100)
    
    'Messagebox
    If MsgBox(msgCdEnd, vbOKCancel, ttlCdEnd) <> vbOK Then
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
    
    If HL_SHP = "co" Then
        'Add callout
        Set shp = ActiveSheet.Shapes.AddShape _
        (msoShapeRoundedRectangularCallout, x1, y1, x2 - x1, y2 - y1)
    
        'Set callout
        With shp
            .Fill.Visible = msoTrue
            .Fill.ForeColor.RGB = RGB(255, 255, 255)
            .Fill.Transparency = 0
            .Fill.Solid
            .Line.Visible = msoTrue
            .Line.ForeColor.RGB = RGB(255, 0, 0)
            .Line.Transparency = 0
            .Line.Weight = 3
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 0, 0)
            .TextFrame2.TextRange.Font.Fill.Visible = msoTrue
            .TextFrame2.TextRange.Font.Fill.Transparency = 0
            .TextFrame2.TextRange.Font.Fill.Solid
            .TextFrame2.TextRange.Font.Bold = msoTrue
        End With
    Else
        'Add rectangle
        Set shp = ActiveSheet.Shapes.AddShape _
        (msoShapeRectangle, x1, y1, x2 - x1, y2 - y1)
    
        'Set rectangle
        With shp
            .Fill.Visible = msoFalse
            .Line.Visible = msoTrue
            .Line.ForeColor.RGB = RGB(255, 0, 0)
            .Line.Transparency = 0
            .Line.Weight = 3
        End With
    End If

SetBefore:
    'Rezoom (To avoid excel bug)
    ActiveWindow.Zoom = zoomed
    ActiveWindow.ScrollColumn = scrolledCol
    ActiveWindow.ScrollRow = scrolledRow
    
    Application.Cursor = xlDefault

End Sub
