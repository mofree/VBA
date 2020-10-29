Sub macro1()
    Dim i, j As Integer
    For i = 1 To 5
        For j = 1 To 5
            If Cells(i, j).Value = 4 Then
                Cells(i, j).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 65535
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If
        Next
    Next
End Sub
