Attribute VB_Name = "Run_analysis_all_sheets"
Sub RunStockAnalysisOnAllWorksheets()

    ' Loop through all worksheets and run StockAnalysis macro on each one
    For Each ws In Worksheets
        ws.Activate
        Call Stockanalysis
    Next ws
    
End Sub
