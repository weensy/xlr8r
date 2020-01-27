VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassCL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Constants by Language

Public Sub SetMsg()
    If LANG = "jp" Then
        msgMulti = "共有ワークブックでは実行できません｡"
        msgOlSc = "重なるショートカットがあります。"
        msgCdStt = "開始点の位置でスペースキーを押してください。"
        msgCdEnd = "終了点の位置でスペースキーを押してください。"
        msgNeCll = "セルの指定が間違っています。"
        msgExRow = "入力値が最大行数を超えています。"
        msgExCol = "入力値が最大列数を超えています。"
        ttlCdStt = "開始点"
        ttlCdEnd = "終了点"
    Else
        msgMulti = "Can't run this add-in on shared workbook."
        msgOlSc = "There are overlapping shortcuts."
        msgCdStt = "Press the spacebar at the start point you want."
        msgCdEnd = "Press the spacebar at the end point you want."
        msgNeCll = "Invalid cell format."
        msgExRow = "Input exceeded maximum row."
        msgExCol = "Input exceeded maximum column."
        ttlCdStt = "Start point"
        ttlCdEnd = "End point"
    End If
End Sub
