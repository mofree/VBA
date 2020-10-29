Sub macro2()
    Dim mySheet As String
    mySheet = ActiveSheet.Name
    Sheets(mySheet).Copy after:=Sheets(mySheet)
    Sheets(mySheet).Activate
End Sub
