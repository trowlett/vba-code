Attribute VB_Name = "Module3"
Sub UnderLine(rw As Integer, cl As Integer)
Attribute UnderLine.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
'    Range("A2:Q2").Select
    Range(Cells(rw, 1), Cells(rw, cl)).Select
'    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
'    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
'    With Selection.Borders(xlEdgeTop)
'        .LineStyle = xlContinuous
'        .ColorIndex = xlAutomatic
'        .TintAndShade = 0
'        .Weight = xlThin
'    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
'    Selection.Borders(xlEdgeRight).LineStyle = xlNone
'    Selection.Borders(xlInsideVertical).LineStyle = xlNone
'    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
 End Sub
