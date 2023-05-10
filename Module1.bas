Attribute VB_Name = "Module1"
Sub StockAnalysis()

    ' Define variables
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim totalVolume As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim lastRow As Long
    Dim summaryRow As Integer
    
    ' Set summary table headers
    summaryRow = 2
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Volume"
    
    ' Loop through all rows in the worksheet
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        
        ' Check if the current row is a new ticker symbol
        If Cells(i, 1).Value <> ticker Then
            
            ' Output yearly change and percent change for previous ticker symbol (if applicable)
            If ticker <> "" Then
                yearlyChange = closingPrice - openingPrice
                If openingPrice <> 0 Then
                    percentChange = yearlyChange / openingPrice
                Else
                    percentChange = 0
                End If
                Range("J" & summaryRow).Value = yearlyChange
                Range("K" & summaryRow).Value = percentChange
                Range("K" & summaryRow).NumberFormat = "0.00%"
                Range("L" & summaryRow).Value = totalVolume
                Range("I" & summaryRow).Value = ticker
                
                ' Highlight yearly change cells in green (positive) or red (negative)
                If yearlyChange > 0 Then
                    Range("J" & summaryRow).Interior.ColorIndex = 4
                ElseIf yearlyChange < 0 Then
                    Range("J" & summaryRow).Interior.ColorIndex = 3
                Else
                    Range("J" & summaryRow).Interior.ColorIndex = 0
                End If
                
                ' Update summary row and reset variables
                summaryRow = summaryRow + 1
                totalVolume = 0
            End If
            
            ' Set opening price and ticker variables for new ticker symbol
            ticker = Cells(i, 1).Value
            openingPrice = Cells(i, 3).Value
            
        End If
        
        ' Add current volume to total volume for current ticker symbol
        totalVolume = totalVolume + Cells(i, 7).Value
        
        ' Set closing price for current ticker symbol
        closingPrice = Cells(i, 6).Value
        
    Next i
    
    ' Output greatest % increase, greatest % decrease, and greatest total volume
    Range("O2").Value = "Greatest % Increase:"
    Range("O3").Value = "Greatest % Decrease:"
    Range("O4").Value = "Greatest Total Volume:"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("Q2").Formula = "=MAX(K:K)"
    Range("Q2").NumberFormat = "0.00%"
    Range("Q3").Formula = "=MIN(K:K)"
    Range("Q3").NumberFormat = "0.00%"
    Range("Q4").Formula = "=MAX(L:L)"
    Range("Q4").NumberFormat = "#,##0"
    Range("P2").Formula = _
        "=INDEX(I:I,MATCH(Q2,K:K,0))"
    Range("P3").Formula = _
        "=INDEX(I:I,MATCH(Q3,K:K,0))"
    Range("P4").Formula = _
        "=INDEX(I:I,MATCH(Q4,L:L,0))"
 
End Sub


Sub SortTickers()
    Range("A1").CurrentRegion.Sort Key1:=Range("A2"), _
        Order1:=xlAscending, Header:=xlYes, MatchCase:=False, _
        Orientation:=xlTopToBottom
End Sub



Sub RunStockAnalysisOnAllWorksheets()

    ' Loop through all worksheets and run StockAnalysis macro on each one
    For Each ws In Worksheets
        ws.Activate
        Call StockAnalysis
    Next ws
    
End Sub

