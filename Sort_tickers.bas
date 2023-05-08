Attribute VB_Name = "Sort_tickers"
Sub SortTickers()
    Range("A1").CurrentRegion.Sort Key1:=Range("A2"), _
        Order1:=xlAscending, Header:=xlYes, MatchCase:=False, _
        Orientation:=xlTopToBottom
End Sub

