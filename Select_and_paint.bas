Attribute VB_Name = "Ä£¿é1"
Sub macro1()
Attribute macro1.VB_ProcData.VB_Invoke_Func = " \n14"
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
