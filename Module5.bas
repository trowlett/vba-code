Attribute VB_Name = "Module5"
Sub tintRowOpen(rw As Integer)
'
    rows(rw).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.79998168889431
        .TintAndShade = 0.9
        .PatternTintAndShade = 0
    End With

End Sub
Sub tintRowAway(rw As Integer)
'
    rows(rw).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.799981688894314
        .TintAndShade = 0.9
        .PatternTintAndShade = 0
    End With
    
End Sub

Sub tintRowHome(rw As Integer)
    rows(rw).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13434879
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End Sub

Sub tintRowClub(rw As Integer)
    rows(rw).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 11796441
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
End Sub
Sub tintRowMISGA(rw As Integer)
    rows(rw).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.79998168889431
        .TintAndShade = 0.9
        .PatternTintAndShade = 0
    End With
End Sub


