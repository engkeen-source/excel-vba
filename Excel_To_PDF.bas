Attribute VB_Name = "Excel_To_PDF"
Sub Export_Sheets_To_PDFs()

    Dim ws As Worksheet
    Dim username As String
    
    'Change This
    username = "tanen"
    
    
    For Each ws In Worksheets
        ws.Select

        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:="C:\Users\" & username & "\Desktop\" & ws.Name & ".pdf", _
        openafterpublish:=False
        
    Next ws

End Sub
